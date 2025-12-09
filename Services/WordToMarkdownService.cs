using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;



namespace WordMarkdownAddIn.Services
{
    public class WordToMarkdownService
    {
        private readonly Application _wordApp;
        private readonly Document _activeDoc;

        public WordToMarkdownService() 
        {
            _wordApp = Globals.ThisAddIn.Application;
            _activeDoc = _wordApp.ActiveDocument;

            if (_activeDoc == null)
                throw new System.Exception("Нет активного документа Word.");
        }
        public List<IWordElement> ExtractDocumentStructure()
        {
            var elements = new List<IWordElement>();

            elements = ExtractParagraphs(elements);
            elements = ExtractTables(elements);


            return elements;

        }
        
        // Вспомогательный метод для извлечения форматирования
        private WordFormattedText ExtractFormattedContent(Range paragraphRange)
        {
            var formattedText = new WordFormattedText();
            var runs = formattedText.Runs;

            // Работаем с Characters для более точного анализа форматирования
            var chars = paragraphRange.Characters;

            if (chars.Count == 0) return formattedText; // Если параграф пустой

            // Начинаем с первого символа
            var firstChar = chars[1]; // Индексация с 1 в COM
            string currentText = firstChar.Text;
            var font = firstChar.Font; // Получаем шрифт первого символа
            var currentRun = new FormattedRun
            {
                Text = currentText,
                IsBold = font.Bold == 1, // Word использует -1 для true, 0 для false
                IsItalic = font.Italic == 1,
                IsUnderline = font.Underline != WdUnderline.wdUnderlineNone,
                IsStrikethrough = font.StrikeThrough != 0, // Зачеркивание
                IsSuperscript = font.Superscript != 0,     // Надстрочный индекс
                IsSubscript = font.Subscript != 0,         // Подстрочный индекс
                SmallCaps = font.SmallCaps != 0,           // Капитель
                AllCaps = font.AllCaps != 0,               // Все заглавные
            };

            // Проходим по остальным символам
            for (int i = 2; i <= chars.Count; i++)
            {
                var charRange = chars[i];
                var newFont = charRange.Font; // Получаем шрифт текущего символа

                // Проверяем, отличается ли форматирование от текущего фрагмента
                if (newFont.Bold == (currentRun.IsBold ? 1 : 0) &&
                    newFont.Italic == (currentRun.IsItalic ? 1 : 0) &&
                    (newFont.Underline != WdUnderline.wdUnderlineNone) == currentRun.IsUnderline &&
                    newFont.StrikeThrough == (currentRun.IsStrikethrough ? 1 : 0) && // Или != 0
                    newFont.Superscript == (currentRun.IsSuperscript ? 1 : 0) &&     // Или != 0
                    newFont.Subscript == (currentRun.IsSubscript ? 1 : 0) &&         // Или != 0
                    newFont.SmallCaps == (currentRun.SmallCaps ? 1 : 0) &&           // Или != 0
                    newFont.AllCaps == (currentRun.AllCaps ? 1 : 0) )                // Или != 0
                {
                    // Формат совпадает, добавляем символ к текущему фрагменту
                    currentRun.Text += charRange.Text;
                }
                else
                {
                    // Формат изменился, сохраняем текущий фрагмент
                    runs.Add(currentRun);

                    // Создаем новый фрагмент с новым форматированием
                    currentRun = new FormattedRun
                    {
                        Text = charRange.Text,
                        IsBold = font.Bold == 1,
                        IsItalic = font.Italic == 1,
                        IsUnderline = font.Underline != WdUnderline.wdUnderlineNone
                        // дописать
                    };
                }
            }

            // Не забываем добавить последний фрагмент
            if (!string.IsNullOrEmpty(currentRun.Text))
            {
                runs.Add(currentRun);
            }

            // Убираем символ конца параграфа из последнего фрагмента, если он там оказался
            if (runs.Count > 0 && runs[runs.Count - 1].Text.EndsWith("\r"))
            {
                runs[runs.Count - 1].Text = runs[runs.Count - 1].Text.TrimEnd('\r');
            }

            return formattedText;
        }

        // 1. Таблицы
        private List<IWordElement> ExtractTables(List<IWordElement> elements) 
        {
            foreach (Table table in _activeDoc.Tables)
            {
                var tableData = new List<List<WordFormattedText>>(); // Изменён тип

                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    var rowData = new List<WordFormattedText>();
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        Range cellRange = table.Cell(i, j).Range;
                        string cellText = cellRange.Text.TrimEnd('\r', '\a'); // Обрезаем специальные символы Word
                        // Создаём экземпляр WordFormattedText из строки
                        var formattedCellText = ExtractFormattedContent(cellRange);
                        rowData.Add(formattedCellText);
                    }
                    tableData.Add(rowData);
                }
               elements.Add(new WordTable(tableData));
            }
            
            return elements;

        }

        // 2. Параграфы
        private List<IWordElement> ExtractParagraphs(List<IWordElement> elements )
        {
            // Обходим все параграфы
            foreach (Paragraph para in _activeDoc.Paragraphs)
            {
                // Убираем символ конца параграфа
                string text = para.Range.Text.TrimEnd('\r', '\a');
                
                if (string.IsNullOrEmpty(text)) continue;
                //Определяем тип параграфа
                string styleName = para.get_Style().NameLocal;

                if (styleName.Contains("Heading") || styleName.Contains("Заголовок")) // Заголовок
                {
                    // Извлекаем форматированный контент
                    var content = ExtractFormattedContent(para.Range);
                    // Создаем WordParagraph с именем стиля и контентом.
                    // Уровень заголовка (HeadingLevel) будет вычислен автоматически при обращении к свойству.
                    elements.Add(new WordParagraph(styleName, content));
                    // Больше ничего делать не нужно, свойство HeadingLevel внутри WordParagraph сделает свою работу.
                }
                else if (styleName == "List Paragraph")
                {
                    // Проверяем тип списка через ListFormat
                    bool isOrdered = (para.Range.ListFormat.ListType != WdListType.wdListBullet);
                    // Это элемент списка
                    var content = ExtractFormattedContent(para.Range);
                    var contentList = new List<WordFormattedText> { content }; // Создаем список и добавляем один элемент
                    elements.Add(new WordListItem(contentList, isOrdered)); // Создаем WordListItem
                }
                else if (styleName == "Quote")
                {
                    // Это цитата
                    var content = ExtractFormattedContent(para.Range);
                    elements.Add(new WordQuote(text, content)); // Создаем WordQuote
                }
                else if (styleName == "Subtitle")
                {
                    // Это подзаголовок
                    var content = ExtractFormattedContent(para.Range);
                    elements.Add(new WordSubtitle(text, content)); // Создаем WordSubtitle
                }
                else if (styleName == "Title")
                {
                    // Это название документа
                    var content = ExtractFormattedContent(para.Range);
                    elements.Add(new WordTitle(text, content)); // Создаем WordTitle
                }
                else 
                {
                    if (!string.IsNullOrEmpty(text)) // Пример фильтрации
                    {
                        var content = ExtractFormattedContent(para.Range);
                        elements.Add(new WordParagraph( para.get_Style().NameLocal, content));
                    }
                }
            }
            return elements;

        }

        // 3. Гиперссылки
        private void ExtractHyperlinks()
        {

        }

        //  4. Изображения
        private void ExtractImages()
        {

        }

        // 5. Закладки
        private void ExtractBookmarks()
        {

        }

        // 6. Сноски
        private void ExtractFootnotes()
        {

        }

    }
}
