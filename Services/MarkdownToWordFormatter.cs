using Markdig;
using Markdig.Parsers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Security.Policy;

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
                .UsePipeTables()
                .Build();
        }

        private WordFormattedText ConvertInlineToWordFormattedText(ContainerInline inline)
        {
            var formattedText = new WordFormattedText();
            //Защита от null
            if (inline == null)
            { 
                return formattedText;
            }

            // Начинаем обход с первого дочернего элемента
            var current = inline.FirstChild;
            while (current != null) 
            {
                // Создаем новый FormattedRun для текущего элемента
                var run = new FormattedRun();

                // Обрабатываем разные типы inline-элементов
                if (current is LiteralInline literal)
                {
                    // Простой текст без форматирования
                    run.Text = literal.Content.ToString();
                    // Все флаги форматирования остаются false(по умолчанию)
                }
                else if (current is EmphasisInline emphasis)
                {
                    // Жирный или курсив текст
                    // Рекурсивно обрабатываем содержимое emphasis
                    var innerText = ConvertInlineToWordFormattedText(emphasis);

                    // Объединяем весь текст из вложенных элементов
                    run.Text = string.Join("",innerText.Runs.Select(r => r.Text));

                    // Определяем тип форматирования по количеству символов
                    // **текст** = 2 символа = жирный
                    // *текст* = 1 символ = курсив
                    run.IsBold = emphasis.DelimiterCount == 2;
                    run.IsItalic = emphasis.DelimiterCount == 1;

                    // Если внутри emphasis есть вложенное форматирование,
                    // можно объединить его с текущим
                    if(innerText.Runs.Count > 0)
                    {
                        // Берем форматирование из первого вложенного элемента
                        var firstInner = innerText.Runs[0];
                        run.IsStrikethrough = firstInner.IsStrikethrough;
                        run.IsUnderline = firstInner.IsUnderline;
                    }
                }
                else if (current is CodeInline code)
                {
                    // Инлайн-код: `код`
                    run.Text += code.Content.ToString();
                    // Можно добавить специальное форматирование для кода
                    // Например, моноширинный шрифт (но это обычно делается на уровне параграфа)
                }
                else if (current is LinkInline link)
                {
                    // Ссылка: [текст] (url)
                    if(link.FirstChild != null)
                    {
                        // Есть текст ссылки - рекурсивно получаем его
                        var linkText = ConvertInlineToWordFormattedText(link);
                        run.Text = string.Join("", linkText.Runs.Select(r => r.Text));
                    }
                    else
                    {
                        // Нет текста - используем URL
                        run.Text = link.Url ?? string.Empty;
                    }

                    // подчеркивание для ссылок
                    run.IsUnderline = true;
                    // URL можно сохранить отдельно, если нужно создать гиперссылку
                    // Но для простоты пока просто текст с подчеркиванием
                }
                else if(current is LineBreakInline)
                {
                    // Перенос строки
                    run.Text = "\r"; // или "\n" в зависимости от системы
                }
                else if (current is HtmlInline html)
                {
                    // HTML-теги в Markdown (например, <br>, <strong>)
                    // Можно обработать специальные теги или просто пропустить
                    run.Text = html.Tag ?? string.Empty;
                }
                else if (current is AutolinkInline autolink)
                {
                    // Автоматическая ссылка (например, https://example.com)
                    run.Text = autolink.Url ?? string.Empty;
                    run.IsUnderline = true; // Ссылка обычно подчеркнута
                }
                else if (current is ContainerInline container)
                {
                    // Если это контейнер (может содержать другие элементы)
                    // Рекурсивно обрабатываем его содержимое
                    var containerText = ConvertInlineToWordFormattedText(container);
                    // Объединяем все runs из контейнера
                    formattedText.Runs.AddRange(containerText.Runs);
                    current = current.NextSibling;
                    continue; // Пропускаем добавление run, т.к. уже добавили через AddRange
                }

                // Добавляем run только если есть текст
                if (!string.IsNullOrEmpty(run.Text))
                {
                    formattedText.Runs.Add(run);
                }

                // Переходим к следующему элементу
                current = current.NextSibling;
            }

            return formattedText;
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
        private IWordElement ProcessHeading(HeadingBlock heading)
        {
            try
            {
                if (heading == null || _activeDoc == null)
                return null;

            
                //1. Извлекаем текст с форматированием
                var formattedText = ConvertInlineToWordFormattedText(heading.Inline);

                // 2. Определяем стиль заголовка
                string styleName = $"Heading {heading.Level}";

                // 3. Создаем WordParagraph
                return new WordParagraph(styleName, formattedText);
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                System.Diagnostics.Debug.WriteLine($"Ошибка при обработке заголовка: {ex.Message}");
                return null;
            }
        }

        //Параграф
        private IWordElement ProcessParagraph(ParagraphBlock paragraph)
        {
            if (paragraph == null || paragraph.Inline.FirstChild == null)
            {
                return null;
            }

            try
            {
                // 1. Преобразуем inline-элементы в WordFormattedText
                var formattedText = ConvertInlineToWordFormattedText(paragraph.Inline);

                // 2. Получаем текущий стиль пользователя
                string currentStyle = GetCurrentParagraphStyle();
                return new WordParagraph(currentStyle, formattedText);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки параграфа: {ex.Message}");
                return null;
            }
        }

        //Блок кода
        private void ProcessCodeBlock(CodeBlock codeBlock)
        {
            if (codeBlock == null || _activeDoc == null)
                return;

            try
            {
                
                // Получаем текст кода из блока
                var sb = new StringBuilder();
           
                string codeText = sb.ToString().TrimEnd('\r', '\n');
                
                if (string.IsNullOrEmpty(codeText))
                    return;

                // Создаем параграф для блока кода
                var codeParagraph = _activeDoc.Content.Paragraphs.Add();
                
                // Устанавливаем моноширинный шрифт
                codeParagraph.Range.Font.Name = "Courier New";   // Возможно изменить
                codeParagraph.Range.Font.Size = 10;
                
                // Вставляем текст кода
                codeParagraph.Range.Text = codeText;
                
                // Применяем стиль для кода (если есть) или оставляем обычный
                // Можно использовать встроенный стиль "Code" если он существует
                try
                {
                    codeParagraph.Range.set_Style("Code");
                }
                catch
                {
                    // Если стиль "Code" не существует, используем обычный стиль
                    codeParagraph.Range.set_Style("Normal");
                }
                
                // Добавляем перенос строки после блока кода
                codeParagraph.Range.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки блока кода: {ex.Message}");
            }
        }

        //Цитата
        private void ProcessQuote(QuoteBlock quoteBlock)
        {
            if (quoteBlock == null || _activeDoc == null)
                return;

            try
            {
                // Обрабатываем все параграфы внутри цитаты
                foreach (var block in quoteBlock)
                {
                    if (block is ParagraphBlock paragraph)
                    {
                        // Создаем параграф для цитаты
                        var quoteParagraph = _activeDoc.Content.Paragraphs.Add();
                        
                        // Извлекаем текст из параграфа цитаты
                        string quoteText = GetTextFormInline(paragraph.Inline);

                        string userStyle = GetCurrentParagraphStyle();

                        if (!string.IsNullOrEmpty(quoteText))
                        {
                            quoteParagraph.Range.Text = quoteText;
                        }
                        
                        // Применяем стиль цитаты
                        try
                        {
                            quoteParagraph.Range.set_Style("Courier New");  // Проверить 
                        }
                        catch
                        {
                            // Если стиль "Quote" не существует, используем обычный стиль
                            quoteParagraph.Range.set_Style(userStyle);   
                            // Добавляем отступ для визуального выделения
                            quoteParagraph.Range.ParagraphFormat.LeftIndent = 36; // 0.5 дюйма
                        }
                        
                        quoteParagraph.Range.InsertParagraphAfter();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки цитаты: {ex.Message}");
            }
        }

        //Таблица - версия с dynamic для совместимости с разными версиями Markdig
        private void ProcessTableDynamic(dynamic tableBlock)
        {
            if (tableBlock == null || _activeDoc == null)
                return;

            try
            {
                // Получаем количество строк и столбцов через dynamic
                int rowCount = tableBlock.Count;
                if (rowCount == 0)
                    return;

                // Определяем количество столбцов из первой строки
                int columnCount = 0;
                dynamic firstRow = tableBlock[0];
                if (firstRow != null)
                {
                    columnCount = firstRow.Count;
                }

                if (columnCount == 0)
                    return;

                // Создаем таблицу в Word
                var range = _activeDoc.Content;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                
                var wordTable = _activeDoc.Tables.Add(
                    range,
                    rowCount,
                    columnCount,
                    WdDefaultTableBehavior.wdWord9TableBehavior,
                    WdAutoFitBehavior.wdAutoFitFixed);

                // Обрабатываем каждую строку таблицы
                for (int rowIndex = 0; rowIndex < rowCount && rowIndex < tableBlock.Count; rowIndex++)
                {
                    dynamic markdownRow = tableBlock[rowIndex];
                    if (markdownRow == null)
                        continue;

                    // Обрабатываем каждую ячейку в строке
                    int rowColumnCount = markdownRow.Count;
                    for (int colIndex = 0; colIndex < columnCount && colIndex < rowColumnCount; colIndex++)
                    {
                        dynamic cell = markdownRow[colIndex];
                        if (cell == null)
                            continue;

                        // Получаем текст ячейки
                        string cellText = string.Empty;
                        int cellCount = cell.Count;
                        if (cellCount > 0)
                        {
                            dynamic firstCellBlock = cell[0];
                            if (firstCellBlock != null)
                            {
                                // Пробуем получить текст через Inline, если это ParagraphBlock
                                try
                                {
                                    if (firstCellBlock is ParagraphBlock cellParagraph)
                                    {
                                        cellText = GetTextFormInline(cellParagraph.Inline);
                                    }
                                    else if (firstCellBlock.Inline != null)
                                    {
                                        // Альтернативный способ получения текста
                                        cellText = GetTextFormInline(firstCellBlock.Inline);
                                    }
                                }
                                catch
                                {
                                    // Если не удалось получить текст, оставляем пустым
                                }
                            }
                        }

                        // Вставляем текст в ячейку Word (индексация с 1)
                        var wordCell = wordTable.Cell(rowIndex + 1, colIndex + 1);
                        wordCell.Range.Text = cellText;
                        
                        // Убираем символ конца параграфа из ячейки
                        wordCell.Range.Text = wordCell.Range.Text.TrimEnd('\r', '\a');
                    }
                }

                // Форматируем таблицу
                wordTable.AutoFormat(
                    WdTableFormat.wdTableFormatGrid1,
                    true,  // Автоподбор размеров
                    false, // Не применять границы
                    false, // Не применять заливку
                    false, // Не применять шрифт
                    false, // Не применять цвет
                    false, // Не применять автоформат
                    false); // Не применять выравнивание

                // Добавляем перенос строки после таблицы
                range.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки таблицы: {ex.Message}");
            }
        }

        //Список
        private void ProcessList(ListBlock listBlock)
        {
            if (listBlock == null || _activeDoc == null)
                return;

            try
            {
                // Определяем тип списка (нумерованный или маркированный)
                bool isOrdered = listBlock.IsOrdered;

                // Обрабатываем каждый элемент списка
                foreach (var item in listBlock)
                {
                    if (item is ListItemBlock listItem)
                    {
                        // Обрабатываем параграфы внутри элемента списка
                        foreach (var block in listItem)
                        {
                            if (block is ParagraphBlock paragraph)
                            {
                                // Создаем параграф для элемента списка
                                var listParagraph = _activeDoc.Content.Paragraphs.Add();
                                
                                // Извлекаем текст элемента списка
                                string itemText = GetTextFormInline(paragraph.Inline);
                                
                                if (!string.IsNullOrEmpty(itemText))
                                {
                                    listParagraph.Range.Text = itemText;
                                }

                                // Применяем форматирование списка
                                if (isOrdered)
                                {
                                    // Нумерованный список
                                    listParagraph.Range.ListFormat.ApplyNumberDefault();
                                }
                                else
                                {
                                    // Маркированный список
                                    listParagraph.Range.ListFormat.ApplyBulletDefault();
                                }

                                listParagraph.Range.InsertParagraphAfter();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки списка: {ex.Message}");
            }
        }

        //Блок кода
        private void ProcessBlock(Block block)
        {
            if (block == null)
                return;

            try
            {
                //Обработать каждый тип блока
                if (block is HeadingBlock heading)
                {
                    ProcessHeading(heading);
                }
                else if (block is ParagraphBlock paragraph)
                {
                    ProcessParagraph(paragraph);
                }
                else if (block is CodeBlock codeBlock)
                {
                    ProcessCodeBlock(codeBlock);
                }
                else if (block is QuoteBlock quoteBlock)
                {
                    ProcessQuote(quoteBlock);
                }
                else if (block.GetType().Name.Contains("Table") && 
                         (block.GetType().Namespace == "Markdig.Extensions.Tables" || 
                          block.GetType().Namespace?.Contains("Markdig.Extensions") == true))
                {
                    // Используем dynamic для обхода проблемы с типами в разных версиях Markdig
                    ProcessTableDynamic(block);
                }
                else if (block is ListBlock listBlock)
                {
                    ProcessList(listBlock);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки блока {block.GetType().Name}: {ex.Message}");
            }
        }

        // с форматированием
        public void ApplyMarkdownToWord(string markdown) 
        {

            // 1. Парсим Markdown через Markdig
            var document = Markdown.Parse(markdown, _pipeline);

            // 2. Преобразуем в коллекцию IWordElement
            var elements = new List<IWordElement>();
            foreach(var block in document)
            {
                var element = ProcessBlock(block);
                if(element != null)
                    elements.Add(element);
            }

            // 3. Применяем все элементы к Word
            foreach(var element in elements)
            {
                element.ApplyToWord(_activeDoc);
            }



        }
        // без форматирования
        public void InsertMarkdownAsPlainText(string markdown) 
        {
        
        }
    }
}
