using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace WordMarkdownAddIn.Services
{
    public interface IWordElement
    {
        string ElementType { get; }
        string ToMarkdown(); // Метод для преобразования элемента в строку Markdown
    }

    public interface IListItemElement : IWordElement { }

    // Класс для представления фрагмента текста с единым форматированием
    public class FormattedRun
    {
        public string ElementType => "Text";
        public string Text { get; set; } = "";      //Сам текст
        public bool IsBold { get; set; }            // Жирный шрифт
        public bool IsItalic { get; set; }          // Курсив
        public bool IsStrikethrough { get; set; }   // Зачеркнутый
        public bool IsUnderline { get; set; }       // Подчеркивание
        public bool IsSuperscript { get; set; }     //Надстрочный текст (например, x²)
        public bool IsSubscript { get; set; }       // Подстрочный текст (например, H₂O)
        public bool SmallCaps { get; set; }         // Уменьшенные заглавные буквы
        public bool AllCaps { get; set; }           // ВСЁ ЗАГЛАВНЫМИ (все буквы заглавные)
    }


    public class WordFormattedText : IWordElement
    {
        public string ElementType => "Content";
        public List<FormattedRun> Runs { get; set; } = new List<FormattedRun>();
        // можно добавить формтирование всего блока (блочное форматирование)
        public string ToMarkdown()
        {
            if (Runs == null || Runs.Count == 0)
                return "";

            var sb = new StringBuilder();
            foreach (var run in Runs)
            {
                if (run == null || string.IsNullOrEmpty(run.Text))
                    continue;

                string text = run.Text;
                string markdown = text;

                // Применяем форматирование в правильном порядке
                // Сначала применяем подстрочный/надстрочный текст к исходному тексту
                if (run.IsSubscript)
                    markdown = $"<sub>{text}</sub>";
                else if (run.IsSuperscript)
                    markdown = $"<sup>{text}</sup>";

                // Затем применяем Markdown форматирование
                if (run.IsStrikethrough)
                    markdown = $"~~{markdown}~~";
                
                if (run.IsBold)
                    markdown = $"**{markdown}**";
                
                if (run.IsItalic)
                    markdown = $"*{markdown}*";

                // Подчеркивание (в Markdown нет нативного синтаксиса, используем HTML)
                // Применяем последним, чтобы обернуть все форматирование
                if (run.IsUnderline)
                    markdown = $"<u>{markdown}</u>";

                sb.Append(markdown);
            }

            return sb.ToString();
        }
    }

    public class WordTable : IWordElement
    {
        public string ElementType => "Table";
        public List<List<WordFormattedText>> Rows { get; set; } = new List<List<WordFormattedText>>();

        public WordTable(List<List<WordFormattedText>> rows)
        {
            Rows = rows ?? new List<List<WordFormattedText>>(); // защита от null
        }
        public string ToMarkdown()
        {
            if (Rows == null || Rows.Count == 0)
                return "";

            var sb = new StringBuilder();
            
            // Обрабатываем первую строку как заголовок таблицы
            if (Rows.Count > 0)
            {
                var headerRow = Rows[0];
                if (headerRow != null && headerRow.Count > 0)
                {
                    // Заголовок таблицы
                    sb.Append("| ");
                    foreach (var cell in headerRow)
                    {
                        string cellContent = cell?.ToMarkdown() ?? "";
                        sb.Append(cellContent.Replace("|", "\\|")); // Экранируем символы |
                        sb.Append(" | ");
                    }
                    sb.AppendLine();

                    // Разделитель заголовка
                    sb.Append("|");
                    for (int i = 0; i < headerRow.Count; i++)
                    {
                        sb.Append(" --- |");
                    }
                    sb.AppendLine();
                }

                // Остальные строки
                for (int i = 1; i < Rows.Count; i++)
                {
                    var row = Rows[i];
                    if (row != null && row.Count > 0)
                    {
                        sb.Append("| ");
                        foreach (var cell in row)
                        {
                            string cellContent = cell?.ToMarkdown() ?? "";
                            sb.Append(cellContent.Replace("|", "\\|")); // Экранируем символы |
                            sb.Append(" | ");
                        }
                        sb.AppendLine();
                    }
                }
            }

            return sb.ToString().TrimEnd();
        }
    }

    public class WordListItem : IWordElement
    {
        public string ElementType => "ListItem";
        public List<WordFormattedText> Contents { get; set; } // Неправильно для одного элемента списка
        public bool IsOrdered { get; set; }

        public WordListItem(List<WordFormattedText> contents, bool isOrdered) 
        {
            Contents = contents;
            IsOrdered = isOrdered;
        }

        public string ToMarkdown() 
        {
            if (Contents == null || Contents.Count == 0)
                return "";

            var sb = new StringBuilder();
            foreach (var content in Contents)
            {
                if (content != null)
                {
                    string markdown = content.ToMarkdown();
                    if (!string.IsNullOrEmpty(markdown))
                    {
                        sb.Append(markdown);
                    }
                }
            }

            return sb.ToString();
        }
    }

    public class WordParagraph : IWordElement
    {
        public string ElementType => "Paragraph";
        public string StyleName { get; set; } = "Normal"; // для заголовков: Heading 1, Normal и т.д.
        public WordFormattedText Content { get; set; } = new WordFormattedText();

        public WordParagraph(string styleName, WordFormattedText content)
        {
            StyleName = styleName;
            Content = content;
        }

        // вычисляемый уровень заголовка
        public int HeadingLevel
        {
            get
            {
                if (StyleName.StartsWith("Heading"))
                {
                    string levelStr = StyleName.Substring("Heading".Length).Trim();
                    if (int.TryParse(levelStr, out int level))
                        return level;
                }
                return 0; // не заголовок
            }
        }

        public string ToMarkdown()
        {
            if (Content == null)
                return "";

            string contentMarkdown = Content.ToMarkdown();
            
            int headingLevel = HeadingLevel;
            if (headingLevel > 0)
            {
                // Преобразуем в заголовок Markdown
                string hashes = new string('#', headingLevel);
                return $"{hashes} {contentMarkdown}";
            }
            else
            {
                // Обычный параграф
                return contentMarkdown;
            }
        }
    }

    public class WordQuote : IWordElement
    {
        public string ElementType => "Quote";
        public string Text { get; set; }
        public WordFormattedText Content { get; set; }

        public WordQuote(string text, WordFormattedText content)
        {
            Text = text;
            Content = content;
        }
        public string ToMarkdown()
        {
            // Используем Content, если он есть, иначе Text
            string quoteText = "";
            if (Content != null)
            {
                quoteText = Content.ToMarkdown();
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                quoteText = Text;
            }

            if (string.IsNullOrEmpty(quoteText))
                return "";

            // Разбиваем на строки и добавляем префикс > для каждой строки
            var lines = quoteText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var sb = new StringBuilder();
            
            foreach (var line in lines)
            {
                sb.Append("> ");
                sb.AppendLine(line);
            }

            return sb.ToString().TrimEnd();
        }

    }

    public class WordSubtitle : IWordElement
    {
        public string ElementType => "Subtitle";
        public string Text { get; set; }
        public WordFormattedText Content { get; set; }

        public WordSubtitle(string text, WordFormattedText content)
        {
            Text = text;
            Content = content;
        }
        public string ToMarkdown()
        {
            // Используем Content, если он есть, иначе Text
            string subtitleText = "";
            if (Content != null)
            {
                subtitleText = Content.ToMarkdown();
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                subtitleText = Text;
            }

            if (string.IsNullOrEmpty(subtitleText))
                return "";

            // Подзаголовок обычно представляется как заголовок уровня 2
            return $"## {subtitleText}";
        }

    }

    public class WordTitle : IWordElement
    {
        public string ElementType => "Title";
        public string Text { get; set; }
        public WordFormattedText Content { get; set; }

        public WordTitle(string text, WordFormattedText content)
        {
            Text = text;
            Content = content;
        }

        public string ToMarkdown()
        {
            // Используем Content, если он есть, иначе Text
            string titleText = "";
            if (Content != null)
            {
                titleText = Content.ToMarkdown();
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                titleText = Text;
            }

            if (string.IsNullOrEmpty(titleText))
                return "";

            // Заголовок обычно представляется как заголовок уровня 1
            return $"# {titleText}";
        }

    }

}
