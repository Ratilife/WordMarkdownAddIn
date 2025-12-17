using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordMarkdownAddIn.Services
{

    public class Token
    {
        public string Text { get; set; }
        public TokenType Type { get; set; }
        public int StartPosition { get; set; }
        public int EndPosition { get; set; }

        // Дополнительно (опционально, для более сложной подсветки):
        public string Language { get; set; }  // язык программирования
        public bool IsHighlighted { get; set; } // нужно ли выделять цветом
    }

    public enum TokenType
    {
        // Базовые типы (уже есть)
        Keyword,        // def, if, return, class
        String,         // "текст", 'текст'
        Comment,        // # комментарий, // комментарий
        Number,         // 123, 3.14
        Operator,       // +, -, =, ==
        Identifier,     // имена переменных/функций (общее)
        Punctuation,    // , . ; ( ) [ ] { }
        Default,        // всё остальное

        // Дополнительные типы для полной подсветки
        Function,       // имена функций (часто выделяются отдельно)
        Class,          // имена классов
        Variable,       // переменные (если нужно отдельно от Identifier)
        Type,           // типы данных (int, string, bool, str)
        Constant,       // константы (PI, MAX_VALUE)
        Builtin,        // встроенные функции (print, len, range в Python)
        Decorator,      // декораторы (@decorator в Python)
        Attribute,      // атрибуты (obj.attr)
        Namespace,      // пространства имён (import module)
        Regex,          // регулярные выражения
        Escape,         // escape-последовательности (\n, \t)
        Tag,            // HTML-теги (<div>, </div>)
        Entity,         // HTML-сущности (&amp;, &lt;)
        Meta,           // метаданные (директивы, препроцессор)
    }

    public class SyntaxHighlighter
    {
        public List<Token> ParseCode(string code, string language)
        {
            return null;
        }
    }
}
