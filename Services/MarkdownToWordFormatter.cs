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
