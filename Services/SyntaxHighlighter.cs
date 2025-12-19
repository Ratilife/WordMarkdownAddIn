using Word = Microsoft.Office.Interop.Word;
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

        private static readonly Dictionary<TokenType, Word.WdColor> TokenColors =
            new Dictionary<TokenType, Word.WdColor>
        {
            [TokenType.Keyword] =       Word.WdColor.wdColorBlue,       // Синий для ключевых слов
            [TokenType.String] =        Word.WdColor.wdColorGreen,      // Зелёный для строк
            [TokenType.Comment] =       Word.WdColor.wdColorGray25,     // Серый для комментариев
            [TokenType.Number] =        Word.WdColor.wdColorOrange,     // Оранжевый для чисел
            [TokenType.Operator] =      Word.WdColor.wdColorDarkRed,    // Тёмно-красный для операторов
            [TokenType.Builtin] =       Word.WdColor.wdColorTeal,       // Бирюзовый для встроенных функций
            [TokenType.Type] =          Word.WdColor.wdColorDarkBlue,    // Тёмно-синий для типов
            [TokenType.Constant] =      Word.WdColor.wdColorViolet,     // Фиолетовый для констант
            [TokenType.Decorator] =     Word.WdColor.wdColorPink,       // Розовый для декораторов
            [TokenType.Regex] =         Word.WdColor.wdColorOliveGreen, // Оливковый для регулярных выражений
            [TokenType.Default] =       Word.WdColor.wdColorAutomatic,  // Автоматический (чёрный) для остального
            [TokenType.Identifier] =    Word.WdColor.wdColorAutomatic,  // Автоматический для идентификаторов
            [TokenType.Punctuation] =   Word.WdColor.wdColorAutomatic   // Автоматический для знаков препинания
    };

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
                TokenType.Operator,
                TokenType.Identifier,   // нет
                TokenType.Punctuation,  // нет
                TokenType.Default,      // нет 
                TokenType.Function,     // нет 
                TokenType.Class,        // нет
                TokenType.Variable      // нет
            };

            // 5. Находим все токены для каждого типа (в порядке приоритета)
            foreach (var tokenType in priorityOrder)
            {
                if (!patterns.ContainsKey(tokenType))
                    continue;

                string pattern = patterns[tokenType];
                try
                {
                    var regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                    var matches = regex.Matches(code);

                    foreach (Match match in matches)
                    {
                        // Проверяем, не перекрывается ли этот токен с уже найденными
                        bool overlaps = tokens.Any(t =>
                            // Новый токен начинается внутри существующего
                            (match.Index >= t.StartPosition && match.Index < t.EndPosition) ||
                            // Новый токен заканчивается внутри существующего
                            (match.Index + match.Length > t.StartPosition && match.Index + match.Length <= t.EndPosition) ||
                            // Новый токен полностью содержит существующий
                            (match.Index <= t.StartPosition && match.Index + match.Length >= t.EndPosition) ||
                            // Существующий токен полностью содержит новый
                            (t.StartPosition <= match.Index && t.EndPosition >= match.Index + match.Length));

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
                catch (ArgumentException ex)
                {
                    // Некорректный паттерн регулярного выражения
                    System.Diagnostics.Debug.WriteLine(
                        $"Ошибка в паттерне для {tokenType} языка {lang}: {ex.Message}");
                    // Продолжаем работу с другими паттернами
                    continue;
                }
                catch (RegexMatchTimeoutException ex)
                {
                    // Паттерн слишком сложный, превышено время выполнения
                    System.Diagnostics.Debug.WriteLine(
                        $"Таймаут при обработке паттерна {tokenType} языка {lang}: {ex.Message}");
                    continue;
                }
            }

            // 6. Сортируем токены по позиции в коде
            tokens = tokens.OrderBy(t => t.StartPosition).ToList();

            // 7. Заполняем пробелы между токенами (текст, который не попал ни под один паттерн)
            // Пробелы и переносы строк не нужно добавлять как отдельные токены,
            // так как они будут сохранять форматирование по умолчанию
            var result = new List<Token>();
            int currentPosition = 0;

            foreach (var token in tokens)
            {
                // Если есть пробел между текущей позицией и началом токена
                if (token.StartPosition > currentPosition)
                {
                    string gapText = code.Substring(currentPosition, token.StartPosition - currentPosition);
                    // Добавляем пробелы только если это не просто пробелы/переносы строк
                    // Пробелы будут сохранять форматирование по умолчанию и не будут перезаписывать цвета
                    if (!string.IsNullOrEmpty(gapText))
                    {
                        // Определяем тип для пробела
                        TokenType gapType = DetermineGapType(gapText);
                        // Добавляем только если это не просто пробелы/переносы строк
                        // Пробелы с типом Default не будут форматироваться, что правильно
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
            // Если это пустая строка - не должно происходить, но на всякий случай
            if (string.IsNullOrEmpty(text))
                return TokenType.Default;

            // Если это ТОЛЬКО пробелы/табы (без переносов строк) - это отступы
            if (Regex.IsMatch(text, @"^[ \t]+$"))
                return TokenType.Default; // или можно создать TokenType.Indent

            // Если это ТОЛЬКО переносы строк
            if (Regex.IsMatch(text, @"^[\r\n]+$"))
                return TokenType.Default; // или TokenType.LineBreak

            // Если это пробелы + переносы строк (комбинация)
            if (Regex.IsMatch(text, @"^[\s]+$"))
                return TokenType.Default;

            // Если есть знаки препинания
            if (Regex.IsMatch(text, @"^[.,;:()\[\]{}]+$"))
                return TokenType.Punctuation;

            // Если это выглядит как идентификатор
            if (Regex.IsMatch(text, @"^\w+$"))
                return TokenType.Identifier;

            // Всё остальное (смешанное содержимое)
            return TokenType.Default;
        }

        /// <summary>
        /// Применяет подсветку синтаксиса к указанному диапазону текста в Word документе
        /// </summary>
        /// <param name="range">Диапазон текста в Word документе</param>
        /// <param name="tokens">Список токенов с информацией о типах и позициях</param>
        /// <param name="originalCode">Исходный код (для проверки соответствия)</param>
        public static void ApplyHighlightingToWordRange(
            Word.Range range,
            List<Token> tokens,
            string originalCode)
        {
            if (range == null || tokens == null || tokens.Count == 0)
                return;

            if (string.IsNullOrEmpty(originalCode))
                return;

            // Получаем текст из Range для проверки соответствия
            string rangeText = range.Text;
            
            // Убираем символ конца параграфа из rangeText для корректного сравнения
            // Word добавляет \r (символ возврата каретки) в конце Range параграфа
            string normalizedRangeText = rangeText.TrimEnd('\r', '\n', '\a');
            
            // Проверяем, что текст Range соответствует originalCode (с учетом возможных различий в переносах строк)
            // Нормализуем оба текста для сравнения
            string normalizedCode = originalCode.Replace("\r\n", "\n").Replace("\r", "\n");
            string normalizedRange = normalizedRangeText.Replace("\r\n", "\n").Replace("\r", "\n");
            
            // ВАЖНО: Используем реальный текст из Range для определения длины
            // Это гарантирует, что мы не выйдем за границы Range
            int maxLength = normalizedRange.Length;
            int rangeStart = range.Start;

            // Если тексты не совпадают, выводим предупреждение, но продолжаем работу
            if (normalizedCode != normalizedRange)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Предупреждение: текст Range не соответствует originalCode. Range length: {normalizedRange.Length}, Code length: {normalizedCode.Length}");
                System.Diagnostics.Debug.WriteLine($"Range text (first 50 chars): {normalizedRange.Substring(0, Math.Min(50, normalizedRange.Length))}");
                System.Diagnostics.Debug.WriteLine($"Code text (first 50 chars): {normalizedCode.Substring(0, Math.Min(50, normalizedCode.Length))}");
            }

            // Применяем форматирование к каждому токену
            foreach (var token in tokens)
            {
                // Проверяем границы токена относительно originalCode
                if (token.StartPosition < 0 || 
                    token.EndPosition <= token.StartPosition ||
                    token.StartPosition >= maxLength ||
                    token.EndPosition > maxLength)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Токен пропущен: позиция [{token.StartPosition}-{token.EndPosition}] выходит за границы (maxLength={maxLength})");
                    continue;
                }

                try
                {
                    // Вычисляем абсолютные позиции в документе
                    // Позиции токенов считаются относительно начала originalCode
                    int tokenStart = rangeStart + token.StartPosition;
                    int tokenEnd = rangeStart + token.EndPosition;

                    // Проверяем, что позиции в пределах Range
                    if (tokenStart < rangeStart || tokenEnd > range.End)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"Токен пропущен: позиции [{tokenStart}-{tokenEnd}] выходят за границы Range [{rangeStart}-{range.End}]");
                        continue;
                    }

                    // Создаём поддиапазон для текущего токена
                    Word.Range tokenRange = range.Duplicate;
                    tokenRange.SetRange(tokenStart, tokenEnd);

                    // Проверяем, что текст токена совпадает с ожидаемым
                    string tokenTextInRange = tokenRange.Text.TrimEnd('\r', '\n', '\a');
                    if (tokenTextInRange != token.Text)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"Предупреждение: текст токена не совпадает. Ожидалось: '{token.Text}', получено: '{tokenTextInRange}'");
                        // Продолжаем работу, так как это может быть из-за форматирования Word
                    }

                    // Применяем цвет в зависимости от типа токена
                    // Пропускаем токены с Default типом, чтобы не перезаписывать форматирование
                    if (token.Type != TokenType.Default && TokenColors.ContainsKey(token.Type))
                    {
                        tokenRange.Font.Color = TokenColors[token.Type];
                        System.Diagnostics.Debug.WriteLine(
                            $"Применен цвет {TokenColors[token.Type]} к токену '{token.Text}' типа {token.Type}");
                    }

                    // Дополнительное форматирование для ключевых слов и встроенных функций
                    if (token.Type == TokenType.Keyword || token.Type == TokenType.Builtin)
                    {
                        tokenRange.Font.Bold = 1; // Жирный шрифт
                    }
                }
                catch (Exception ex)
                {
                    // Игнорируем ошибки форматирования отдельных токенов
                    System.Diagnostics.Debug.WriteLine(
                        $"Ошибка форматирования токена [{token.StartPosition}-{token.EndPosition}]: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Применяет подсветку синтаксиса к выделенному тексту в Word документе
        /// </summary>
        /// <param name="app">Приложение Word</param>
        /// <param name="language">Язык программирования (python, csharp, java, javascript, 1c)</param>
        public static void HighlightSelectedCode(Word.Application app, string language = "python")
        {
            if (app == null || app.Selection == null)
                return;

            Word.Selection selection = app.Selection;
            string code = selection.Text;

            if (string.IsNullOrEmpty(code))
                return;

            // 1. Парсим код на токены
            List<Token> tokens = ParseCode(code, language);

            // 2. Применяем форматирование к выделенному тексту
            ApplyHighlightingToWordRange(selection.Range, tokens, code);
        }

        /// <summary>
        /// Применяет подсветку синтаксиса к блоку кода в Word документе
        /// </summary>
        /// <param name="range">Диапазон текста с кодом в Word документе</param>
        /// <param name="code">Текст кода для парсинга</param>
        /// <param name="language">Язык программирования (python, csharp, java, javascript, 1c)</param>
        public static void HighlightCodeBlock(Word.Range range, string code, string language = "python")
        {
            if (range == null || string.IsNullOrEmpty(code))
            {
                System.Diagnostics.Debug.WriteLine("HighlightCodeBlock: range или code пусты");
                return;
            }

            try
            {
                // Нормализуем язык
                string normalizedLanguage = (language ?? "").ToLower().Trim();
                if (string.IsNullOrEmpty(normalizedLanguage))
                {
                    normalizedLanguage = "python";
                }

                System.Diagnostics.Debug.WriteLine($"HighlightCodeBlock: начинаем подсветку. Язык: '{normalizedLanguage}', Длина кода: {code.Length}");
                
                // 1. Парсим код на токены
                List<Token> tokens = ParseCode(code, normalizedLanguage);
                
                System.Diagnostics.Debug.WriteLine($"HighlightCodeBlock: найдено токенов: {tokens.Count}");
                
                if (tokens.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("HighlightCodeBlock: токены не найдены. Проверьте, поддерживается ли язык.");
                    return;
                }
                
                // 2. Применяем форматирование к Range
                ApplyHighlightingToWordRange(range, tokens, code);
                
                System.Diagnostics.Debug.WriteLine("HighlightCodeBlock: подсветка применена успешно");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка подсветки синтаксиса: {ex.Message}\n{ex.StackTrace}");
            }
        }

    }
}
