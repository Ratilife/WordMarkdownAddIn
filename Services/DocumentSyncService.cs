using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordMarkdownAddIn.Services
{
    public static class DocumentSyncService
    {
        /// <summary>
        /// Пространство имён XML, используемое для элементов, связанных с исходным Markdown-контентом.
        /// </summary>
        public const string NamespaceUri = "urn:markdown/source";    // определяет уникальный идентификатор для разметки, связанной с Markdown-контентом в документе Word.

        /// <summary>
        /// Загружает Markdown-содержимое из XML-части активного документа Word.
        /// </summary>
        /// <param name="app">Объект приложения Microsoft Word.</param>
        /// <returns>Строка с Markdown-содержимым или <c>null</c>, если документ не активен, 
        /// XML-часть не найдена или не содержит данных.</returns>
        public static string LoadMarkdownFromActiveDocument(Word.Application app)
        {
            //Проверить - доступен ли Word и активный документ
            //Найти - XML-часть с Markdown-данными
            //Извлечь - содержимое из элемента <content>
            //Вернуть - Markdown-текст или null если не найдено
            if (app == null || app.ActiveDocument == null) return null;                                             // убеждается, что объект Word и активный документ существуют
            var doc = app.ActiveDocument;                                                                           // сохраняет активный документ в локальную переменную doc
            Office.CustomXMLPart part = FindExistingPart(doc);                                                      // вызывает метод FindExistingPart для поиска части с Markdown            
            if (part == null) return null;                                                                          // если XML-часть не найдена - возвращает null, Значит, в документе ранее не сохранялся Markdown                                                                        
            try
            {
                var node = part.SelectSingleNode("/*[local-name()='markdown']/*[local-name()='content']");          // ищет конкретный XML-узел
                                                                                                                    // /*[local-name()='markdown'] - корневой элемент с локальным именем 'markdown' (игнорирует пространство имен)
                                                                                                                    // /*[local-name()='content'] - дочерний элемент 'content'
                                                                                                                    // local-name() позволяет найти элементы независимо от префиксов
                if (node != null)                                                                                   // если XPath-запрос вернул узел
                {
                    return node.Text;                                                                               // возвращает текстовое содержимое узла
                }                                                                                                   // Для CDATA-секций Text возвращает содержимое без CDATA-обертки
                                                                                                                    // Если содержимое содержало "]]>", оно было разбито на несколько CDATA секций при сохранении
                                                                                                                    // XML парсер автоматически объединяет несколько CDATA секций обратно в одно содержимое
            }
            catch { }                                                                                               //  игнорирует любые ошибки при работе с XML
            return null;
        }


        /// <summary>
        /// Сохраняет Markdown-содержимое в активный документ Word как пользовательскую XML-часть.
        /// Если в документе уже существует сохранённая версия Markdown, она удаляется перед добавлением новой.
        /// </summary>
        /// <param name="app">Объект приложения Microsoft Word.</param>
        /// <param name="markdown">Строка с Markdown-содержимым для сохранения.</param>
        public static void SaveMarkdownToActiveDocument(Word.Application app, string markdown)   
        {
            //Проверить - есть ли доступный документ Word
            //Найти - существующую версию Markdown (если есть)
            //Удалить - старую версию чтобы избежать дублирования
            //Построить - новую XML-структуру с актуальным Markdown
            //Сохранить - XML в документ как пользовательскую часть
            if (app == null || app.ActiveDocument == null) return;                                                  // убеждается, что объект Word и активный документ существуют
            var doc = app.ActiveDocument;                                                                           // сохраняет активный документ в локальную переменную для удобства
            var existing = FindExistingPart(doc);                                                                   // вызывает вспомогательный метод для поиска уже сохраненного Markdown
                                                                                                                    // FindExistingPart ищет XML-часть с пространством имен urn:markdown/source
            if (existing != null)
            {
                existing.Delete();                                                                                  // если Markdown уже сохранялся ранее - удаляет старую XML-часть
            }
            var xml = BuildXml(markdown ?? string.Empty);                                                           // вызывает метод для создания XML-структуры
            doc.CustomXMLParts.Add(xml);                                                                            // добавляет подготовленный XML в пользовательские части документа

        }


        /// <summary>
        /// Создаёт XML-строку с Markdown-содержимым, обёрнутым в CDATA.
        /// Правильно обрабатывает случаи, когда содержимое содержит "]]>", разбивая CDATA на несколько секций.
        /// </summary>
        /// <param name="content">Markdown-содержимое, которое нужно включить в XML.</param>
        /// <returns>Строка в формате XML с корневым элементом 'markdown' и дочерним элементом 'content',
        /// где содержимое заключено в секцию CDATA (или несколько секций, если содержимое содержит "]]>").</returns>
        private static string BuildXml(string content)
        {
            // Обернуть markdown в CDATA внутри корневого пространства имен
            // Если содержимое содержит "]]>", нужно разбить CDATA на несколько секций
            // Заменяем "]]>" на "]]]]><![CDATA[>", что создаст две CDATA секции с правильным содержимым
            var safeContent = content.Replace("]]>", "]]]]><![CDATA[>");                                            // заменяет "]]>" на "]]]]><![CDATA[>"
                                                                                                                    // Это разбивает CDATA на две секции: <![CDATA[...]]]><![CDATA[>...]]>
                                                                                                                    // В результате XML парсер увидит: <![CDATA[...]]]><![CDATA[>...]]>
                                                                                                                    // А фактическое содержимое останется: ...]]>...
            return "<md:markdown xmlns:md='" + NamespaceUri + "'>" +                                                // создает элемент <md:markdown>
                                                                                                                    // добавляет атрибут xmlns:md='urn:markdown/source'
                                                                                                                    // md: - префикс, связанный с нашим пространством имен
                                                                                                                    // NamespaceUri - константа urn:markdown/source
                "<md:content><![CDATA[" + safeContent + "]]></md:content>" +                                        // создает элемент <md:content>
                                                                                                                    // оборачивает content в <![CDATA[ ... ]]>
                                                                                                                    // CDATA предотвращает интерпретацию специальных XML-символов
                                                                                                                    // Если content содержал "]]>", он будет разбит на несколько CDATA секций
                "</md:markdown>";                                                                                   // завершает корневой элемент                                                                                     
        }


        /// <summary>
        /// Находит существующую пользовательскую XML-часть в документе, соответствующую пространству имён Markdown.
        /// </summary>
        /// <param name="doc">Документ Word, в котором выполняется поиск.</param>
        /// <returns>Объект <see cref="Office.CustomXMLPart"/>, если часть найдена; в противном случае — <c>null</c>.</returns>
        private static Office.CustomXMLPart FindExistingPart(Word.Document doc)
        {
            // Получить все пользовательские XML-части документа
            // Перебрать каждую часть в цикле
            // Для каждой части:
            //  Получить корневой элемент
            //  Проверить его пространство имен
            //  Если совпадает с нашим - вернуть часть
            // Если не найдено - вернуть null
            try
            {
                Office.CustomXMLParts parts = doc.CustomXMLParts;                                                   // получает все пользовательские XML-части документа
                foreach (Office.CustomXMLPart p in parts)                                                           // перебирает все XML-части документа одна за другой; p - текущая проверяемая XML-часть
                {
                    try
                    {
                        var root = p.DocumentElement;                                                               // получает корневой XML-элемент текущей части
                                                                                                                    // DocumentElement - свойство, возвращающее root element XML-документа
                        if (root != null && string.Equals(root.NamespaceURI, NamespaceUri, StringComparison.OrdinalIgnoreCase))  // root != null - корневой элемент существует
                                                                                                                                 // string.Equals(...) - пространство имен корневого элемента совпадает с нашим (urn:markdown/source)                    
                                                                                                                                 // StringComparison.OrdinalIgnoreCase - сравнение без учета регистра 
                        {
                            return p;                                                                                            //  возвращает найденную XML-часть   
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }
    }
}
