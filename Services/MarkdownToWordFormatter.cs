using Markdig;
using Markdig.Parsers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordMarkdownAddIn.Services
{
    public class MarkdownToWordFormatter
    {
        private readonly Application _wordApp;
        private readonly Document _activeDoc;
        private readonly MarkdownPipeline _pipeline;

        public MarkdownToWordFormatter() 
        {
            _wordApp = Globals.ThisAddIn.Application;
            _activeDoc = _wordApp.ActiveDocument;
            _pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .Build();
        }

        private void ProcessMarkdownDocument(Markdig.Syntax.MarkdownDocument doc) 
        {
            //Обойти все дочерние узлы
            foreach (var block in doc) 
            {
                ProcessBlock(block);
            }
            // ... другие типы
        }

        private string GetTextFormInline(Markdig.Syntax.Inlines.ContainerInline inline) 
        {
            if (inline == null) return string.Empty;

            var sb = new StringBuilder();

            //Обходим все дочерние элементы
            var current = inline.FirstChild;
            while (current != null) 
            {
                if (current is LiteralInline literal)
                {
                    //Простой текст - просто добавляем
                    sb.Append(literal.Content.ToString());
                }
                else if (current is EmphasisInline emphasis)
                {
                    // Жирный или курсив - рекурсивнополучает текст внутри
                    sb.Append(GetTextFormInline(emphasis));
                }
                else if (current is LinkInline link)
                {
                    //Ссылка - берем текст ссылки (или URL)
                    if (link.FirstChild != null)
                    {
                        sb.Append(GetTextFormInline(link));
                    }
                    else
                    {
                        sb.Append(link.Url ?? string.Empty);
                    }

                }
                else if (current is CodeInline code) 
                {
                    // Инлайн-код
                    sb.Append(code.Content.ToString());
                }
                current = current.NextSibling;
                
            }
            return sb.ToString();
        }

        private void ApplyInlineToWordRange(ContainerInline inline, Range wordRange)
        {
            if (inline == null || wordRange == null)
                return;

            var current = inline.FirstChild;
            while (current != null)
            {
                if (current is LiteralInline literal)
                {
                    // Простой текст - вставляем как есть
                    wordRange.InsertAfter(literal.Content.ToString());
                    // Перемещаем курсор в конец вставленного текста
                    wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
                else if (current is EmphasisInline emphasis)
                {
                    // Жирный или курсив
                    // Сохраняем текущее форматирование
                    var originalBold = wordRange.Font.Bold;
                    var originalItalic = wordRange.Font.Italic;

                    // Определяем тип форматирования
                    bool isBold = emphasis.DelimiterCount == 2;  // **текст** - жирный
                    bool isItalic = emphasis.DelimiterCount == 1; // *текст* - курсив

                    // Применяем форматирование
                    if (isBold)
                        wordRange.Font.Bold = 1;
                    else if (isItalic)
                        wordRange.Font.Italic = 1;

                    // Рекурсивно обрабатываем содержимое emphasis
                    ApplyInlineToWordRange(emphasis, wordRange);

                    // Восстанавливаем форматирование
                    wordRange.Font.Bold = originalBold;
                    wordRange.Font.Italic = originalItalic;
                }
                else if (current is LinkInline link)
                {
                    // Ссылка
                    string linkText = GetTextFormInline(link);
                    if (string.IsNullOrEmpty(linkText))
                        linkText = link.Url ?? string.Empty;

                    // Вставляем текст ссылки
                    wordRange.InsertAfter(linkText);
                    wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // Создаем гиперссылку (опционально)
                    if (!string.IsNullOrEmpty(link.Url))
                    {
                        // Можно создать гиперссылку через Hyperlinks.Add()
                        // Но это сложнее, пока просто вставляем текст
                    }
                }
                else if (current is CodeInline code)
                {
                    // Инлайн-код - обычно моноширинный шрифт
                    var originalFont = wordRange.Font.Name;
                    wordRange.Font.Name = "Courier New"; // или другой моноширинный

                    wordRange.InsertAfter(code.Content.ToString());
                    wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    wordRange.Font.Name = originalFont;
                }
                else if (current is LineBreakInline)
                {
                    // Перенос строки
                    wordRange.InsertAfter("\r");
                    wordRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                // Переходим к следующему элементу
                current = current.NextSibling;
            }
        }

        private string GetCurrentParagraphStyle()
        {
            try
            {
                // Получаем стиль из позиции курсора
                if (_wordApp?.Selection?.Range != null)
                {
                    string styleName = _wordApp.Selection.Range.get_Style().NameLocal;
                    // Проверяем, что это стиль параграфа, а не символов
                    if (!string.IsNullOrEmpty(styleName) && !styleName.StartsWith("Стиль символов"))
                    {
                        return styleName;
                    }
                }
            }
            catch
            {
                // Игнорируем ошибки
            }

            // Возвращаем "Normal" по умолчанию
            return "Normal";
        }

        //Заголовок
        private void ProcessHeading(HeadingBlock heading)
        {
            if (heading == null || _activeDoc == null)
                return;

            try
            {
                // 1. Извлекаем текст заголовка
                string headingText = GetTextFormInline(heading.Inline);

                if (string.IsNullOrEmpty(headingText))
                    return; // Пустой заголовок - пропускаем

                // 2. Создаем параграф в Word
                var paragraph = _activeDoc.Content.Paragraphs.Add();

                // 3. Вставляем текст
                paragraph.Range.Text = headingText;

                // 4. Применяем стиль заголовка
                string styleName = $"Heading {heading.Level}";
                paragraph.Range.set_Style(styleName);

                // 5. Добавляем перенос строки после заголовка
                paragraph.Range.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                System.Diagnostics.Debug.WriteLine($"Ошибка при обработке заголовка: {ex.Message}");
            }
        }

        //Параграф
        private void ProcessParagraph(ParagraphBlock paragraph)
        {
            if (paragraph == null || paragraph.Inline.FirstChild == null)
            {
                return;
            }

            try
            {
                var wordParagraph = _activeDoc.Content.Paragraphs.Add();
                //Проверяем есть ли содержание
                if (paragraph.Inline == null || paragraph.Inline.NextSibling == null)
                {
                    //Пусстой параграф - создаем пустую строку
                    string userStyle = GetCurrentParagraphStyle();
                    wordParagraph.Range.set_Style(userStyle);
                    wordParagraph.Range.InsertParagraphAfter();
                    return;
                }
               
                // Извлекаем текст параграфа
                string paragraphText = GetTextFormInline(paragraph.Inline);

                if (string.IsNullOrEmpty(paragraphText))
                {
                    // Пустой параграф
                    var emptyPara = _activeDoc.Content.Paragraphs.Add();
                    string userStyle = GetCurrentParagraphStyle();
                    emptyPara.Range.set_Style(userStyle);
                    emptyPara.Range.InsertParagraphAfter();
                    return;
                }

                // Создаем параграф в Word
                var newParagraph = _activeDoc.Content.Paragraphs.Add();

                // Вставляем текст
                newParagraph.Range.Text = paragraphText;

                // Получаем стиль, выбранный пользователем в позиции курсора
                string currentStyle = GetCurrentParagraphStyle();

                // Применяем стиль пользователя (или "Normal" по умолчанию)
                newParagraph.Range.set_Style(currentStyle);

                // Добавляем перенос строки после параграфа
                newParagraph.Range.InsertParagraphAfter();

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки параграфа: {ex.Message}");
            }
        }

        private void ProcessBlock(Block block)
        {
            //Обработать каждый тип блока
            if (block is HeadingBlock heading) 
            {
                
            }
        }

        // с форматированием
        public void ApplyMarkdownToWord(string markdown) 
        {

            // 1. Распарсить
            var document = Markdown.Parse(markdown, _pipeline);
            
            // 2. Обойти дерево
            ProcessMarkdownDocument(document);


            
        }
        // без форматирования
        public void InsertMarkdownAsPlainText(string markdown) 
        {
        
        }
    }
}
