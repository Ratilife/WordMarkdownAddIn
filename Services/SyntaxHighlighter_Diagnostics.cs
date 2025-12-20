using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordMarkdownAddIn.Services
{
    /// <summary>
    /// Временный класс для диагностики проблем с позициями токенов
    /// Используйте этот метод для проверки проблемы #6
    /// </summary>
    public static class SyntaxHighlighterDiagnostics
    {
        /// <summary>
        /// Детальная диагностика проблемы с позициями токенов
        /// Вызывает этот метод вместо HighlightCodeBlock для проверки
        /// </summary>
        public static void DiagnoseTokenPositions(
            Word.Range range,
            string code,
            string language = "python")
        {
            if (range == null || string.IsNullOrEmpty(code))
            {
                System.Diagnostics.Debug.WriteLine("=== ДИАГНОСТИКА: range или code пусты ===");
                return;
            }

            System.Diagnostics.Debug.WriteLine("=== НАЧАЛО ДИАГНОСТИКИ ПОЗИЦИЙ ТОКЕНОВ ===");
            System.Diagnostics.Debug.WriteLine($"Язык: {language}");
            System.Diagnostics.Debug.WriteLine($"Длина originalCode: {code.Length}");
            
            // 1. Получаем текст из Range
            string rangeText = range.Text;
            string normalizedRangeText = rangeText.TrimEnd('\r', '\n', '\a');
            
            System.Diagnostics.Debug.WriteLine($"Длина rangeText (до нормализации): {rangeText.Length}");
            System.Diagnostics.Debug.WriteLine($"Длина normalizedRangeText: {normalizedRangeText.Length}");
            System.Diagnostics.Debug.WriteLine($"Range.Start: {range.Start}, Range.End: {range.End}");
            
            // 2. Нормализуем оба текста
            string normalizedCode = code.Replace("\r\n", "\n").Replace("\r", "\n");
            string normalizedRange = normalizedRangeText.Replace("\r\n", "\n").Replace("\r", "\n");
            
            System.Diagnostics.Debug.WriteLine($"Длина normalizedCode: {normalizedCode.Length}");
            System.Diagnostics.Debug.WriteLine($"Длина normalizedRange: {normalizedRange.Length}");
            
            // 3. Сравниваем тексты побайтово
            bool textsMatch = normalizedCode == normalizedRange;
            System.Diagnostics.Debug.WriteLine($"Тексты совпадают: {textsMatch}");
            
            if (!textsMatch)
            {
                System.Diagnostics.Debug.WriteLine("=== РАЗЛИЧИЯ В ТЕКСТАХ ===");
                
                // Показываем первые 100 символов каждого текста
                int previewLength = Math.Min(100, Math.Max(normalizedCode.Length, normalizedRange.Length));
                System.Diagnostics.Debug.WriteLine($"OriginalCode (первые {previewLength} символов):");
                System.Diagnostics.Debug.WriteLine(GetTextPreview(normalizedCode, previewLength));
                System.Diagnostics.Debug.WriteLine($"RangeText (первые {previewLength} символов):");
                System.Diagnostics.Debug.WriteLine(GetTextPreview(normalizedRange, previewLength));
                
                // Находим первое различие
                int firstDifference = FindFirstDifference(normalizedCode, normalizedRange);
                if (firstDifference >= 0)
                {
                    System.Diagnostics.Debug.WriteLine($"Первое различие на позиции: {firstDifference}");
                    int contextStart = Math.Max(0, firstDifference - 20);
                    int contextLength = Math.Min(40, normalizedCode.Length - contextStart);
                    System.Diagnostics.Debug.WriteLine($"Контекст originalCode: '{GetTextPreview(normalizedCode.Substring(contextStart, contextLength), contextLength)}'");
                    if (contextStart < normalizedRange.Length)
                    {
                        int rangeContextLength = Math.Min(contextLength, normalizedRange.Length - contextStart);
                        System.Diagnostics.Debug.WriteLine($"Контекст rangeText: '{GetTextPreview(normalizedRange.Substring(contextStart, rangeContextLength), rangeContextLength)}'");
                    }
                }
            }
            
            // 4. Парсим токены
            List<Token> tokens = SyntaxHighlighter.ParseCode(code, language);
            System.Diagnostics.Debug.WriteLine($"=== ПАРСИНГ ТОКЕНОВ ===");
            System.Diagnostics.Debug.WriteLine($"Найдено токенов: {tokens.Count}");
            
            if (tokens.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("ВНИМАНИЕ: Токены не найдены! Проверьте язык и паттерны.");
                System.Diagnostics.Debug.WriteLine("=== КОНЕЦ ДИАГНОСТИКИ ===");
                return;
            }
            
            // 5. Анализируем каждый токен
            int maxLength = normalizedRange.Length;
            int rangeStart = range.Start;
            int skippedTokens = 0;
            int appliedTokens = 0;
            
            System.Diagnostics.Debug.WriteLine($"maxLength (длина normalizedRange): {maxLength}");
            System.Diagnostics.Debug.WriteLine($"rangeStart: {rangeStart}");
            System.Diagnostics.Debug.WriteLine($"=== АНАЛИЗ ТОКЕНОВ ===");
            
            foreach (var token in tokens)
            {
                System.Diagnostics.Debug.WriteLine($"--- Токен: '{token.Text}' (тип: {token.Type}) ---");
                System.Diagnostics.Debug.WriteLine($"  Позиция в originalCode: [{token.StartPosition}-{token.EndPosition}]");
                System.Diagnostics.Debug.WriteLine($"  Длина токена: {token.Text.Length}");
                
                // Проверка границ относительно originalCode
                bool withinOriginalBounds = token.StartPosition >= 0 && 
                                           token.EndPosition <= code.Length &&
                                           token.EndPosition > token.StartPosition;
                System.Diagnostics.Debug.WriteLine($"  В границах originalCode: {withinOriginalBounds}");
                
                // Проверка границ относительно normalizedRange
                bool withinRangeBounds = token.StartPosition >= 0 && 
                                       token.EndPosition <= maxLength &&
                                       token.EndPosition > token.StartPosition;
                System.Diagnostics.Debug.WriteLine($"  В границах normalizedRange (maxLength={maxLength}): {withinRangeBounds}");
                
                if (!withinRangeBounds)
                {
                    System.Diagnostics.Debug.WriteLine($"  ❌ ТОКЕН БУДЕТ ПРОПУЩЕН: выходит за границы");
                    skippedTokens++;
                    continue;
                }
                
                // Вычисляем абсолютные позиции
                int tokenStart = rangeStart + token.StartPosition;
                int tokenEnd = rangeStart + token.EndPosition;
                System.Diagnostics.Debug.WriteLine($"  Абсолютные позиции в документе: [{tokenStart}-{tokenEnd}]");
                System.Diagnostics.Debug.WriteLine($"  Границы Range: [{rangeStart}-{range.End}]");
                
                // Проверка абсолютных позиций
                bool withinDocumentBounds = tokenStart >= rangeStart && tokenEnd <= range.End;
                System.Diagnostics.Debug.WriteLine($"  В границах Range документа: {withinDocumentBounds}");
                
                if (!withinDocumentBounds)
                {
                    System.Diagnostics.Debug.WriteLine($"  ❌ ТОКЕН БУДЕТ ПРОПУЩЕН: выходит за границы Range");
                    skippedTokens++;
                    continue;
                }
                
                // Проверяем, что текст токена соответствует ожидаемому в Range
                try
                {
                    Word.Range tokenRange = range.Duplicate;
                    tokenRange.SetRange(tokenStart, tokenEnd);
                    string tokenTextInRange = tokenRange.Text.TrimEnd('\r', '\n', '\a');
                    
                    System.Diagnostics.Debug.WriteLine($"  Текст в Range: '{tokenTextInRange}'");
                    System.Diagnostics.Debug.WriteLine($"  Ожидаемый текст: '{token.Text}'");
                    System.Diagnostics.Debug.WriteLine($"  Тексты совпадают: {tokenTextInRange == token.Text}");
                    
                    if (tokenTextInRange != token.Text)
                    {
                        System.Diagnostics.Debug.WriteLine($"  ⚠️ ВНИМАНИЕ: Текст токена не совпадает!");
                        System.Diagnostics.Debug.WriteLine($"    Длина в Range: {tokenTextInRange.Length}, ожидаемая: {token.Text.Length}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"  ✅ Токен корректен, будет применен");
                        appliedTokens++;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"  ❌ ОШИБКА при создании Range для токена: {ex.Message}");
                    skippedTokens++;
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"=== ИТОГИ ===");
            System.Diagnostics.Debug.WriteLine($"Всего токенов: {tokens.Count}");
            System.Diagnostics.Debug.WriteLine($"Применено: {appliedTokens}");
            System.Diagnostics.Debug.WriteLine($"Пропущено: {skippedTokens}");
            System.Diagnostics.Debug.WriteLine($"Тексты совпадают: {textsMatch}");
            
            if (!textsMatch)
            {
                System.Diagnostics.Debug.WriteLine("⚠️ КРИТИЧЕСКАЯ ПРОБЛЕМА: Тексты не совпадают!");
                System.Diagnostics.Debug.WriteLine("   Это основная причина проблемы с позициями токенов.");
                System.Diagnostics.Debug.WriteLine("   Токены парсятся относительно originalCode,");
                System.Diagnostics.Debug.WriteLine("   но применяются к range, где текст может отличаться.");
            }
            
            if (skippedTokens > 0)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ ПРОБЛЕМА: {skippedTokens} токенов пропущено из-за неверных позиций");
            }
            
            System.Diagnostics.Debug.WriteLine("=== КОНЕЦ ДИАГНОСТИКИ ===");
        }
        
        private static string GetTextPreview(string text, int length)
        {
            if (string.IsNullOrEmpty(text))
                return "(пусто)";
            
            string preview = text.Substring(0, Math.Min(length, text.Length));
            // Заменяем невидимые символы на видимые
            preview = preview.Replace("\r", "\\r")
                            .Replace("\n", "\\n")
                            .Replace("\t", "\\t")
                            .Replace(" ", "·"); // точка для пробелов
            return preview;
        }
        
        private static int FindFirstDifference(string text1, string text2)
        {
            int minLength = Math.Min(text1.Length, text2.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (text1[i] != text2[i])
                    return i;
            }
            
            if (text1.Length != text2.Length)
                return minLength;
            
            return -1; // Тексты идентичны
        }
    }
}

