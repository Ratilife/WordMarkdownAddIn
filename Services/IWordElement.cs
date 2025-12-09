using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace WordMarkdownAddIn.Services
{
    public interface IWordElement
    {
        string ElementType { get; }
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

    }

    public class WordTable : IWordElement
    {
        public string ElementType => "Table";
        public List<List<WordFormattedText>> Rows { get; set; } = new List<List<WordFormattedText>>();

        public WordTable(List<List<WordFormattedText>> rows)
        {
            Rows = rows ?? new List<List<WordFormattedText>>(); // защита от null
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
        
    }

}
