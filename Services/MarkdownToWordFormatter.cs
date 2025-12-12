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
