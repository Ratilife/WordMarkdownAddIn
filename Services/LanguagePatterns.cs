using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordMarkdownAddIn.Services
{
    public class LanguagePatterns
    {
        // Статический словарь со всеми паттернами
        public static readonly Dictionary<string, Dictionary<TokenType, string>> Patterns =
            new Dictionary<string, Dictionary<TokenType, string>>
            {
                ["python"] = GetPythonPatterns(),
                ["csharp"] = GetCSharpPatterns(),
                ["java"] = GetJavaPatterns(),
                ["javascript"] = GetJavaScriptPatterns(),
                ["1c"] = Get1CPatterns()
            };

        private static Dictionary<TokenType, string> GetPythonPatterns()
        {
            return new Dictionary<TokenType, string>
            {
                [TokenType.Comment] = @"#.*",
                [TokenType.String] = @"("""".*?""""|'''.*?'''|""[^""]*""|'[^']*')",
                [TokenType.Keyword] = @"\b(def|if|elif|else|for|while|return|class|import|from|as|try|except|finally|with|pass|break|continue|yield|lambda|and|or|not|in|is|None|True|False|assert|del|global|nonlocal|raise|async|await)\b",
                [TokenType.Builtin] = @"\b(print|len|range|str|int|float|list|dict|tuple|set|open|input|type|isinstance|hasattr|getattr|setattr|delattr|__init__|__str__|__repr__)\b",
                [TokenType.Number] = @"\b\d+\.?\d*([eE][+-]?\d+)?\b",
                [TokenType.Operator] = @"(==|!=|<=|>=|<<|>>|\*\*|//|&&|\|\||[+\-*/%=<>!&|^~])",
                [TokenType.Decorator] = @"@\w+"
            };
        }

        private static Dictionary<TokenType, string> GetCSharpPatterns()
        {
            return new Dictionary<TokenType, string>
            {
                [TokenType.Comment] = @"(//.*|/\*[\s\S]*?\*/)",
                [TokenType.String] = @"(@""[^""]*""|""[^""]*"")",
                [TokenType.Keyword] = @"\b(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|foreach|goto|if|implicit|in|int|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|virtual|void|volatile|while|async|await|yield|var|dynamic|get|set|value|partial|where|select|from|let|join|into|orderby|group|by|ascending|descending)\b",
                [TokenType.Type] = @"\b(int|string|bool|char|byte|sbyte|short|ushort|uint|long|ulong|float|double|decimal|object|dynamic|var)\b",
                [TokenType.Number] = @"\b\d+\.?\d*([fFdDmMlL])?\b",
                [TokenType.Operator] = @"(==|!=|<=|>=|&&|\|\||[+\-*/%=<>!&|^~?:])"
            };
        }

        private static Dictionary<TokenType, string> GetJavaPatterns()
        {
            return new Dictionary<TokenType, string>
            {
                [TokenType.Comment] = @"(//.*|/\*[\s\S]*?\*/)",
                [TokenType.String] = @"(""[^""]*"")",
                [TokenType.Keyword] = @"\b(abstract|assert|boolean|break|byte|case|catch|char|class|const|continue|default|do|double|else|enum|extends|final|finally|float|for|goto|if|implements|import|instanceof|int|interface|long|native|new|package|private|protected|public|return|short|static|strictfp|super|switch|synchronized|this|throw|throws|transient|try|void|volatile|while|true|false|null)\b",
                [TokenType.Type] = @"\b(boolean|byte|char|short|int|long|float|double|void|String|Object|Integer|Double|Float|Boolean|Character|Byte|Short|Long)\b",
                [TokenType.Number] = @"\b\d+\.?\d*([fFdDlL])?\b",
                [TokenType.Operator] = @"(==|!=|<=|>=|&&|\|\||[+\-*/%=<>!&|^~?:])",
                [TokenType.Decorator] = @"@\w+"
            };
        }

        private static Dictionary<TokenType, string> GetJavaScriptPatterns()
        {
            return new Dictionary<TokenType, string>
            {
                [TokenType.Comment] = @"(//.*|/\*[\s\S]*?\*/)",
                [TokenType.String] = @"(`[^`]*`|""[^""]*""|'[^']*')",
                [TokenType.Keyword] = @"\b(function|var|let|const|if|else|for|while|do|switch|case|break|continue|return|try|catch|finally|throw|new|this|super|extends|class|import|export|async|await|yield|default|from|as|of|in|instanceof|typeof|void|delete|with|debugger|true|false|null|undefined|Infinity|NaN)\b",
                [TokenType.Builtin] = @"\b(console|document|window|Array|Object|String|Number|Boolean|Date|Math|JSON|Promise|Set|Map|WeakSet|WeakMap|RegExp|Error|TypeError|ReferenceError|parseInt|parseFloat|isNaN|isFinite|encodeURI|decodeURI|encodeURIComponent|decodeURIComponent)\b",
                [TokenType.Number] = @"\b(0[xX][0-9a-fA-F]+|0[oO][0-7]+|0[bB][01]+|\d+\.?\d*([eE][+-]?\d+)?)\b",
                [TokenType.Operator] = @"(===|!==|==|!=|<=|>=|&&|\|\||[+\-*/%=<>!&|^~?:])",
                [TokenType.Regex] = @"/(?:[^/\\]|\\.)+/[gimuy]*"
            };
        }

        private static Dictionary<TokenType, string> Get1CPatterns()
        {
            return new Dictionary<TokenType, string>
            {
                [TokenType.Comment] = @"//.*",
                [TokenType.String] = @"""([^""]|"""")*""",
                [TokenType.Keyword] = @"\b(Процедура|Функция|КонецПроцедуры|КонецФункции|Если|Тогда|Иначе|ИначеЕсли|КонецЕсли|Пока|Цикл|КонецЦикла|Для|По|Каждого|Из|Попытка|Исключение|КонецПопытки|ВызватьИсключение|Возврат|Продолжить|Прервать|Перейти|Перем|Выполнить|Новый|СоздатьОбъект|Найти|НайтиСтроки|Сообщить|Истина|Ложь|Неопределено|Null|ПустаяСтрока|Не|И|Или)\b",
                [TokenType.Number] = @"\b\d+\.?\d*\b",
                [TokenType.Operator] = @"(==|<>|<=|>=|[+\-*/%=<>!])",
                [TokenType.Constant] = @"\b(Истина|Ложь|Неопределено|Null|ПустаяСтрока)\b"
            };
        }
    }
}
