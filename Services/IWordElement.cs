using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordMarkdownAddIn.Services
{
    public interface IWordElement
    {
        string ElementType { get; }
    }

    public class WordFormattedText : IWordElement
    {
        public string ElementType => "Text";
        public string Text { get; set; } = "";
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsStrikethrough { get; set; }
        public string StyleName { get; set; } = ""; // для заголовков: Heading 1, Normal и т.д.
    }

    public class WordTable : IWordElement
    {
        public string ElementType => "Table";
        public List<List<WordFormattedText>> Rows { get; set; } = new List<List<WordFormattedText>>();
    }

    public class WordListItem : IWordElement
    {
        public string ElementType => "List";
        public List<WordFormattedText> Items { get; set; } = new List<WordFormattedText>();
        public bool IsOrdered { get; set; }
    }

    public class WordParagraph : IWordElement
    {
        public string ElementType => "Paragraph";
        public string Text { get; set; } = "";
        public string StyleName { get; set; } = ""; // для заголовков: Heading 1, Normal и т.д.
    }
}
