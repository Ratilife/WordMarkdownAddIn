using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        public Token(string text, TokenType type, int startPosition, int endPosition)
        {
            Text = text;
            Type = type;
            StartPosition = startPosition;
            EndPosition = endPosition;
        }

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

        // Статические словари с паттернами (как показывал ранее)
        private static readonly Dictionary<string, Dictionary<TokenType, string>> LanguagePatterns =
            WordMarkdownAddIn.Services.LanguagePatterns.Patterns;


        public static List<Token> ParseCode(string code, string language)
        {

            if (string.IsNullOrEmpty(code))
                return new List<Token>();

            // 1. Нормализуем язык (приводим к нижнему регистру)
            string lang = (language ?? "").ToLower().Trim();

            // Если язык не поддерживается, возвращаем весь код как Default
            if (!LanguagePatterns.ContainsKey(lang))
            {
                return new List<Token>
                {
                    new Token(code, TokenType.Default, 0, code.Length)
                };
            }

            // 2. Получаем паттерны для языка
            var patterns = LanguagePatterns[lang];

            // 3. Список для хранения всех найденных токенов
            var tokens = new List<Token>();

            // 4. Порядок применения паттернов (важно!)
            // Сначала комментарии и строки (чтобы не парсить код внутри них),
            // потом ключевые слова, числа, операторы, остальное
            var priorityOrder = new[]
            {
                TokenType.Comment,
                TokenType.String,
                TokenType.Regex,
                TokenType.Decorator,
                TokenType.Keyword,
                TokenType.Builtin,
                TokenType.Type,
                TokenType.Constant,
                TokenType.Number,
                TokenType.Operator
            };

            // 5. Находим все токены для каждого типа (в порядке приоритета)
            foreach (var tokenType in priorityOrder)
            {
                if (!patterns.ContainsKey(tokenType))
                    continue;

                string pattern = patterns[tokenType];
                var regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                var matches = regex.Matches(code);

                foreach (Match match in matches)
                {
                    // Проверяем, не перекрывается ли этот токен с уже найденными
                    bool overlaps = tokens.Any(t =>
                        (match.Index >= t.StartPosition && match.Index < t.EndPosition) ||
                        (t.StartPosition >= match.Index && t.StartPosition < match.Index + match.Length));

                    if (!overlaps)
                    {
                        tokens.Add(new Token(
                            match.Value,
                            tokenType,
                            match.Index,
                            match.Index + match.Length
                        ));
                    }
                }
            }

            // 6. Сортируем токены по позиции в коде
            tokens = tokens.OrderBy(t => t.StartPosition).ToList();

            // 7. Заполняем пробелы между токенами (текст, который не попал ни под один паттерн)
            var result = new List<Token>();
            int currentPosition = 0;

            foreach (var token in tokens)
            {
                // Если есть пробел между текущей позицией и началом токена
                if (token.StartPosition > currentPosition)
                {
                    string gapText = code.Substring(currentPosition, token.StartPosition - currentPosition);
                    if (!string.IsNullOrWhiteSpace(gapText) || gapText.Contains("\n") || gapText.Contains("\r"))
                    {
                        // Определяем тип для пробела
                        TokenType gapType = DetermineGapType(gapText);
                        result.Add(new Token(gapText, gapType, currentPosition, token.StartPosition));
                    }
                }

                result.Add(token);
                currentPosition = token.EndPosition;
            }

            // Добавляем оставшийся текст в конце
            if (currentPosition < code.Length)
            {
                string remainingText = code.Substring(currentPosition);
                if (!string.IsNullOrEmpty(remainingText))
                {
                    TokenType remainingType = DetermineGapType(remainingText);
                    result.Add(new Token(remainingText, remainingType, currentPosition, code.Length));
                }
            }

            return result;
        }

        // Вспомогательный метод для определения типа пробела между токенами
        private static TokenType DetermineGapType(string text)
        {
            // Если это только пробелы/табы/переносы строк - Default
            if (string.IsNullOrWhiteSpace(text))
                return TokenType.Default;

            // Если есть знаки препинания - Punctuation
            if (Regex.IsMatch(text, @"^[.,;:()\[\]{}]+$"))
                return TokenType.Punctuation;

            // Если это выглядит как идентификатор (имя переменной/функции)
            if (Regex.IsMatch(text, @"^\w+$"))
                return TokenType.Identifier;

            // Всё остальное - Default
            return TokenType.Default;
        }
    }
}
