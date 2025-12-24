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
using System.Security.Cryptography.X509Certificates;

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
                System.Diagnostics.Debug.WriteLine("[ConvertInlineToWordFormattedText] inline равен null");
                return formattedText;
            }

            System.Diagnostics.Debug.WriteLine($"[ConvertInlineToWordFormattedText] Начало обработки inline, FirstChild: {(inline.FirstChild != null ? inline.FirstChild.GetType().Name : "null")}");

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

                    // Определяем тип форматирования по количеству символов
                    
                    
                    foreach (var innerRun in innerText.Runs)
                    {
                        if (emphasis.DelimiterCount == 2)
                            if (emphasis.DelimiterChar == '*')
                            {
                                // **текст** = жирный
                                innerRun.IsBold = true;
                            }
                            else if (emphasis.DelimiterChar == '~')
                            {
                                // ~~текст~~ = зачёркнутый
                                innerRun.IsStrikethrough = true;
                            }

                        if (emphasis.DelimiterCount == 1)
                            // *текст* = 1 символ = курсив
                            innerRun.IsItalic = true;
                        if(emphasis.DelimiterCount == 3)
                        {
                            // ***текст *** =3 жирный + курсив 
                            innerRun.IsBold = true;
                            innerRun.IsItalic = true;
                        }
                    }

                    formattedText.Runs.AddRange(innerText.Runs);
                    current = current.NextSibling;
                    continue;
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

            System.Diagnostics.Debug.WriteLine($"[ConvertInlineToWordFormattedText] Завершено, создано Runs: {formattedText.Runs.Count}");
            if (formattedText.Runs.Count > 0)
            {
                string totalText = string.Join("", formattedText.Runs.Select(r => r?.Text ?? ""));
                System.Diagnostics.Debug.WriteLine($"[ConvertInlineToWordFormattedText] Общий текст: '{totalText}'");
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
                {
                    System.Diagnostics.Debug.WriteLine("[ProcessHeading] heading или _activeDoc равен null");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"[ProcessHeading] Обработка заголовка уровня {heading.Level}");
                System.Diagnostics.Debug.WriteLine($"[ProcessHeading] heading.Inline: {(heading.Inline != null ? "не null" : "null")}");
                System.Diagnostics.Debug.WriteLine($"[ProcessHeading] heading.Inline.FirstChild: {(heading.Inline?.FirstChild != null ? heading.Inline.FirstChild.GetType().Name : "null")}");

                //1. Извлекаем текст с форматированием
                if (heading.Inline == null)
                {
                    System.Diagnostics.Debug.WriteLine("[ProcessHeading] ВНИМАНИЕ: heading.Inline равен null!");
                    return null;
                }

                var formattedText = ConvertInlineToWordFormattedText(heading.Inline);

                // Проверяем, есть ли текст в заголовке
                string headingText = "";
                if (formattedText != null && formattedText.Runs != null && formattedText.Runs.Count > 0)
                {
                    headingText = string.Join("", formattedText.Runs.Select(r => r?.Text ?? ""));
                    System.Diagnostics.Debug.WriteLine($"[ProcessHeading] Текст заголовка: '{headingText}' (длина: {headingText.Length})");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[ProcessHeading] ВНИМАНИЕ: formattedText пустой или нет Runs!");
                }

                // 2. Создаем WordTitle с уровнем заголовка
                var wordTitle = new WordTitle("", formattedText, heading.Level);
                System.Diagnostics.Debug.WriteLine($"[ProcessHeading] WordTitle создан успешно, уровень: {heading.Level}");
                return wordTitle;
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                System.Diagnostics.Debug.WriteLine($"[ProcessHeading] Ошибка при обработке заголовка: {ex.Message}\n{ex.StackTrace}");
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
        private IWordElement ProcessCodeBlock(CodeBlock codeBlock)
        {
            if (codeBlock == null)
                return null;

            try
            {
                // 1. Извлекаем текст кода из блока
                var sb = new StringBuilder();
                
                // CodeBlock содержит строки кода в свойстве Lines
                if (codeBlock.Lines.Count > 0)
                {
                    foreach (var line in codeBlock.Lines)
                    {
                        if (line != null)
                        {
                            sb.AppendLine(line.ToString());
                        }
                    }
                }
                
                string codeText = sb.ToString().TrimEnd('\r', '\n');
                
                if (string.IsNullOrEmpty(codeText))
                    return null;

                // 2. Извлекаем язык программирования (если есть)
                string language = string.Empty;
                if (codeBlock is FencedCodeBlock fencedCodeBlock)
                {
                    // FencedCodeBlock имеет информацию о языке
                    language = fencedCodeBlock.Info ?? string.Empty;
                }

                // 3. Создаем WordCodeBlock
                return new WordCodeBlock(codeText, language);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки блока кода: {ex.Message}");
                return null;
            }
        }

        //Цитата
        private IWordElement ProcessQuote(QuoteBlock quoteBlock)
        {
            if (quoteBlock == null)
                return null;

            try
            {
                // 1. Создаем WordFormattedText для объединения всех параграфов цитаты
                var combinedFormattedText = new WordFormattedText();

                // 2. Обрабатываем все параграфы внутри цитаты
                foreach (var block in quoteBlock)
                {
                    if (block is ParagraphBlock paragraph)
                    {
                        // Извлекаем форматированный текст из параграфа цитаты
                        var paragraphText = ConvertInlineToWordFormattedText(paragraph.Inline);
                        
                        // Объединяем runs из параграфа с общим содержимым цитаты
                        if (paragraphText != null && paragraphText.Runs.Count > 0)
                        {
                            combinedFormattedText.Runs.AddRange(paragraphText.Runs);
                            
                            // Добавляем перенос строки между параграфами (если нужно)
                            // Можно добавить специальный run с переносом строки
                        }
                    }
                }

                // 3. Если цитата пустая, возвращаем null
                if (combinedFormattedText.Runs.Count == 0)
                    return null;

                // 4. Создаем WordQuote
                return new WordQuote("", combinedFormattedText);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки цитаты: {ex.Message}");
                return null;
            }
        }

        //Таблица - версия с dynamic для совместимости с разными версиями Markdig
        private IWordElement ProcessTableDynamic(dynamic tableBlock)
        {
            if (tableBlock == null)
                return null;

            try
            {
                // 1. Определяем размеры таблицы
                int rowCount = tableBlock.Count;
                if (rowCount == 0)
                    return null;

                // Определяем количество столбцов из первой строки
                int columnCount = 0;
                dynamic firstRow = tableBlock[0];
                if (firstRow != null)
                {
                    columnCount = firstRow.Count;
                }

                if (columnCount == 0)
                    return null;

                // 2. Создаем структуру данных для WordTable
                var rows = new List<List<WordFormattedText>>();

                // 3. Обрабатываем каждую строку таблицы
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    dynamic markdownRow = tableBlock[rowIndex];
                    if (markdownRow == null)
                        continue;

                    // Создаем список ячеек для текущей строки
                    var row = new List<WordFormattedText>();

                    // 4. Обрабатываем каждую ячейку в строке
                    for (int colIndex = 0; colIndex < columnCount && colIndex < markdownRow.Count; colIndex++)
                    {
                        dynamic cell = markdownRow[colIndex];
                        
                        // Извлекаем содержимое ячейки с форматированием
                        WordFormattedText cellContent = ExtractCellContent(cell);
                        row.Add(cellContent);
                    }

                    // Добавляем строку в таблицу
                    rows.Add(row);
                }

                // 5. Создаем объект WordTable (НЕ применяем к Word!)
                return new WordTable(rows);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки таблицы: {ex.Message}");
                return null;
            }
        }

        // Вспомогательный метод для извлечения содержимого ячейки
        private WordFormattedText ExtractCellContent(dynamic cell)
        {
            var formattedText = new WordFormattedText();

            if (cell == null)
                return formattedText;

            // Ячейка может содержать несколько блоков (ParagraphBlock и т.д.)
            int cellCount = cell.Count;
            if (cellCount == 0)
                return formattedText;

            // Обрабатываем все блоки в ячейке
            for (int i = 0; i < cellCount; i++)
            {
                dynamic cellBlock = cell[i];
                if (cellBlock == null)
                    continue;

                // Если это ParagraphBlock, извлекаем inline-элементы
                if (cellBlock is ParagraphBlock cellParagraph)
                {
                    // Используем ConvertInlineToWordFormattedText для получения форматирования
                    var paraText = ConvertInlineToWordFormattedText(cellParagraph.Inline);
                    
                    // Объединяем runs из параграфа с общим содержимым ячейки
                    formattedText.Runs.AddRange(paraText.Runs);
                }
                else if (cellBlock.Inline != null)
                {
                    // Альтернативный способ - если блок имеет Inline напрямую
                    var blockText = ConvertInlineToWordFormattedText(cellBlock.Inline);
                    formattedText.Runs.AddRange(blockText.Runs);
                }
            }

            return formattedText;
        }

        //Список
        // Возвращает null, так как список обрабатывается отдельно в ApplyMarkdownToWord
        // из-за того, что список может содержать несколько элементов
        private IWordElement ProcessList(ListBlock listBlock)
        {
            // Список обрабатывается отдельно в ApplyMarkdownToWord,
            // так как он может содержать несколько элементов
            // Возвращаем null, чтобы указать, что список нужно обработать отдельно
            return null;
        }

        // Вспомогательный метод для обработки списка (вызывается из ApplyMarkdownToWord)
        private List<IWordElement> ProcessListItems(ListBlock listBlock)
        {
            var elements = new List<IWordElement>();

            if (listBlock == null)
                return elements;

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
                        var itemContents = new List<WordFormattedText>();

                        foreach (var block in listItem)
                        {
                            if (block is ParagraphBlock paragraph)
                            {
                                // Извлекаем форматированный текст элемента списка
                                var itemFormattedText = ConvertInlineToWordFormattedText(paragraph.Inline);
                                
                                if (itemFormattedText != null && itemFormattedText.Runs.Count > 0)
                                {
                                    itemContents.Add(itemFormattedText);
                                }
                            }
                        }

                        // Создаем WordListItem для каждого элемента списка
                        if (itemContents.Count > 0)
                        {
                            var listItemElement = new WordListItem(itemContents, isOrdered);
                            elements.Add(listItemElement);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обработки списка: {ex.Message}");
            }

            return elements;
        }

        //Блок кода
        private IWordElement ProcessBlock(Block block)
        {
            if (block == null)
            {
                System.Diagnostics.Debug.WriteLine("[ProcessBlock] block равен null");
                return null;
            }

            try
            {
                string blockType = block.GetType().Name;
                System.Diagnostics.Debug.WriteLine($"[ProcessBlock] Обработка блока типа: {blockType}");

                //Обработать каждый тип блока и вернуть результат
                if (block is HeadingBlock heading)
                {
                    System.Diagnostics.Debug.WriteLine($"[ProcessBlock] Найден HeadingBlock, уровень: {heading.Level}");
                    var result = ProcessHeading(heading);
                    System.Diagnostics.Debug.WriteLine($"[ProcessBlock] ProcessHeading вернул: {(result != null ? result.GetType().Name : "null")}");
                    return result;
                }
                else if (block is ParagraphBlock paragraph)
                {
                    return ProcessParagraph(paragraph);
                }
                else if (block is CodeBlock codeBlock)
                {
                    return ProcessCodeBlock(codeBlock);
                }
                else if (block is QuoteBlock quoteBlock)
                {
                    return ProcessQuote(quoteBlock);
                }
                else if (block is ThematicBreakBlock thematicBreak)
                {
                    return ProcessThematicBreak(thematicBreak);
                }
                else if (block.GetType().Name.Contains("Table") && 
                         (block.GetType().Namespace == "Markdig.Extensions.Tables" || 
                          block.GetType().Namespace?.Contains("Markdig.Extensions") == true))
                {
                    // Используем dynamic для обхода проблемы с типами в разных версиях Markdig
                    return ProcessTableDynamic(block);
                }
                else if (block is ListBlock listBlock)
                {
                    return ProcessList(listBlock);
                }
                
                // Если тип блока не распознан, возвращаем null
                System.Diagnostics.Debug.WriteLine($"[ProcessBlock] Неизвестный тип блока: {blockType}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProcessBlock] Ошибка обработки блока {block.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                return null;
            }
        }

        //Разделитель (ThematicBreak)
        private IWordElement ProcessThematicBreak(ThematicBreakBlock thematicBreak)
        {
            if (thematicBreak == null)
                return null;

            try
            {
                // Создаем пустой параграф для разделителя
                return new WordEmptyParagraph();
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                System.Diagnostics.Debug.WriteLine($"Ошибка при обработке разделителя: {ex.Message}");
                return null;
            }
        }

        // с форматированием
        public void ApplyMarkdownToWord(string markdown) 
        {
            System.Diagnostics.Debug.WriteLine("=== [ApplyMarkdownToWord] НАЧАЛО ===");
            System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Длина markdown: {markdown?.Length ?? 0}");
            if (markdown != null)
            {
                int length = Math.Min(500, markdown.Length);
                System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Первые {length} символов markdown:\n{markdown.Substring(0, length)}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[ApplyMarkdownToWord] markdown равен null");
            }

            // 1. Парсим Markdown через Markdig
            var document = Markdown.Parse(markdown, _pipeline);
            System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Markdown распарсен, количество блоков: {document.Count}");

            // Очищаем документ перед вставкой нового контента
            _activeDoc.Content.Delete();

            // Удаляем оставшийся пустой параграф после Delete()
            // Word всегда оставляет один пустой параграф, его нужно удалить
            if (_activeDoc.Content.Paragraphs.Count > 0)
            {
                var lastParagraph = _activeDoc.Content.Paragraphs[_activeDoc.Content.Paragraphs.Count];
                // Проверяем, что параграф действительно пустой (только символ конца параграфа)
                string paragraphText = lastParagraph.Range.Text.TrimEnd('\r', '\a');
                if (string.IsNullOrEmpty(paragraphText))
                {
                    lastParagraph.Range.Delete();
                }
            }

            // 2. Преобразуем в коллекцию IWordElement
            var elements = new List<IWordElement>();
            int blockIndex = 0;
            foreach(var block in document)
            {
                blockIndex++;
                string blockType = block?.GetType().Name ?? "null";
                System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Блок #{blockIndex}: {blockType}");

                // Списки обрабатываем отдельно, так как они могут содержать несколько элементов
                if (block is ListBlock listBlock)
                {
                    var listItems = ProcessListItems(listBlock);
                    System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Добавлено элементов списка: {listItems.Count}");
                    elements.AddRange(listItems);
                }
                else
                {
                    IWordElement element = ProcessBlock(block);
                    if(element != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Элемент добавлен: {element.GetType().Name}, ElementType: {element.ElementType}");
                        elements.Add(element);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] ВНИМАНИЕ: ProcessBlock вернул null для блока типа {blockType}");
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Всего элементов создано: {elements.Count}");

            // 3. Применяем все элементы к Word
            int elementIndex = 0;
            foreach(var element in elements)
            {
                elementIndex++;
                System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Применение элемента #{elementIndex}: {element.GetType().Name}, ElementType: {element.ElementType}");
                try
                {
                    element.ApplyToWord(_activeDoc);
                    System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Элемент #{elementIndex} успешно применен");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] ОШИБКА при применении элемента #{elementIndex}: {ex.Message}\n{ex.StackTrace}");
                }
            }
        }
        // без форматирования
        public void InsertMarkdownAsPlainText(string markdown) 
        {
        
        }
    }
}
