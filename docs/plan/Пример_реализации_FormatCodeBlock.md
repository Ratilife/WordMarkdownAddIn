# Пример реализации FormatCodeBlock с удалением маркеров ```

## 📋 Объяснение проблемы

Текущий код просто заменяет весь текст на `element.Content`, но это не всегда работает правильно, потому что:
1. `element.Content` может содержать маркеры ``` если они не были правильно извлечены
2. Нужно удалить маркеры из самого документа, а не просто заменить текст

## 💡 Решение: Пошаговое удаление маркеров

Вместо замены всего текста, мы будем:
1. Применить форматирование к блоку кода
2. Удалить маркеры ``` из начала блока
3. Удалить маркеры ``` из конца блока
4. Повторно применить форматирование

## 🔧 Полный исправленный метод

```csharp
/// <summary>
/// Применение форматирования для блока кода
/// </summary>
public void FormatCodeBlock(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        // ШАГ 1: Вычисляем позиции блока кода в документе
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range codeRange = _activeDoc.Range(start, end);

        // ШАГ 2: Применяем форматирование к блоку кода
        codeRange.Font.Name = "Consolas";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        codeRange.ParagraphFormat.LeftIndent = 18;
        codeRange.ParagraphFormat.RightIndent = 18;
        codeRange.ParagraphFormat.SpaceBefore = 6;
        codeRange.ParagraphFormat.SpaceAfter = 6;

        // ШАГ 3: Получаем текущий текст блока кода
        string currentText = codeRange.Text;
        
        if (string.IsNullOrEmpty(currentText))
            return;

        // ШАГ 4: Удаляем маркеры ``` из НАЧАЛА блока кода
        // Блок кода может начинаться так: ```csharp\n или ```\n
        if (currentText.StartsWith("```"))
        {
            // Ищем конец первой строки (где заканчивается ```язык или просто ```)
            // Это может быть \n (один символ) или \r\n (два символа)
            int firstLineEnd = currentText.IndexOf('\n');
            
            if (firstLineEnd == -1)
            {
                // Если нет перевода строки, ищем \r
                firstLineEnd = currentText.IndexOf('\r');
            }

            if (firstLineEnd > 0)
            {
                // Нашли перевод строки - удаляем всю первую строку (```язык\n или ```\n)
                // Вычисляем длину: от начала до символа после \n
                int removeLength = firstLineEnd + 1; // +1 чтобы удалить и сам \n
                
                // Проверяем, может быть \r\n (два символа)
                if (firstLineEnd > 0 && currentText[firstLineEnd] == '\r' && 
                    firstLineEnd + 1 < currentText.Length && currentText[firstLineEnd + 1] == '\n')
                {
                    removeLength = firstLineEnd + 2; // +2 для \r\n
                }

                // Создаем диапазон для удаления (от начала до конца первой строки)
                Range startRange = _activeDoc.Range(codeRange.Start, codeRange.Start + removeLength);
                startRange.Delete();
                
                // ВАЖНО: После удаления нужно обновить диапазон codeRange
                // Потому что позиции сдвинулись
                codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - removeLength);
            }
            else
            {
                // Нет перевода строки - удаляем только ``` (3 символа)
                Range startRange = _activeDoc.Range(codeRange.Start, codeRange.Start + 3);
                startRange.Delete();
                codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - 3);
            }
        }

        // ШАГ 5: Удаляем маркеры ``` из КОНЦА блока кода
        // Получаем обновленный текст после удаления начала
        string updatedText = codeRange.Text;
        
        if (!string.IsNullOrEmpty(updatedText))
        {
            // Проверяем разные варианты окончания блока кода
            if (updatedText.EndsWith("\r\n```"))
            {
                // Вариант 1: \r\n``` (5 символов: \r + \n + ```)
                Range endRange = _activeDoc.Range(codeRange.End - 5, codeRange.End);
                endRange.Delete();
            }
            else if (updatedText.EndsWith("\n```"))
            {
                // Вариант 2: \n``` (4 символа: \n + ```)
                Range endRange = _activeDoc.Range(codeRange.End - 4, codeRange.End);
                endRange.Delete();
            }
            else if (updatedText.EndsWith("```"))
            {
                // Вариант 3: просто ``` (3 символа)
                Range endRange = _activeDoc.Range(codeRange.End - 3, codeRange.End);
                endRange.Delete();
            }
        }

        // ШАГ 6: Повторно применяем форматирование после удаления маркеров
        // (на случай, если форматирование сбросилось)
        codeRange.Font.Name = "Consolas";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        codeRange.ParagraphFormat.LeftIndent = 18;
        codeRange.ParagraphFormat.RightIndent = 18;
        codeRange.ParagraphFormat.SpaceBefore = 6;
        codeRange.ParagraphFormat.SpaceAfter = 6;

        Debug.WriteLine($"[FormatCodeBlock] Применено форматирование блока кода, удалены маркеры ```");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatCodeBlock] Ошибка: {ex.Message}");
    }
}
```

## 📝 Упрощенная версия (если сложно)

Если логика выше кажется сложной, вот упрощенная версия:

```csharp
/// <summary>
/// Применение форматирования для блока кода (упрощенная версия)
/// </summary>
public void FormatCodeBlock(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range codeRange = _activeDoc.Range(start, end);

        // Применяем форматирование
        codeRange.Font.Name = "Consolas";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        codeRange.ParagraphFormat.LeftIndent = 18;
        codeRange.ParagraphFormat.RightIndent = 18;
        codeRange.ParagraphFormat.SpaceBefore = 6;
        codeRange.ParagraphFormat.SpaceAfter = 6;

        // Удаляем маркеры из начала
        string text = codeRange.Text;
        if (text.StartsWith("```"))
        {
            // Ищем первую строку (до \n)
            int newlinePos = text.IndexOf('\n');
            if (newlinePos > 0)
            {
                // Удаляем первую строку (```язык\n)
                Range toDelete = _activeDoc.Range(codeRange.Start, codeRange.Start + newlinePos + 1);
                toDelete.Delete();
                // Обновляем диапазон
                codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - (newlinePos + 1));
            }
            else
            {
                // Нет \n, удаляем только ```
                Range toDelete = _activeDoc.Range(codeRange.Start, codeRange.Start + 3);
                toDelete.Delete();
                codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - 3);
            }
        }

        // Удаляем маркеры из конца
        text = codeRange.Text;
        if (text.EndsWith("```"))
        {
            Range toDelete = _activeDoc.Range(codeRange.End - 3, codeRange.End);
            toDelete.Delete();
        }
        else if (text.EndsWith("\n```"))
        {
            Range toDelete = _activeDoc.Range(codeRange.End - 4, codeRange.End);
            toDelete.Delete();
        }
        else if (text.EndsWith("\r\n```"))
        {
            Range toDelete = _activeDoc.Range(codeRange.End - 5, codeRange.End);
            toDelete.Delete();
        }

        // Повторно применяем форматирование
        codeRange.Font.Name = "Consolas";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        codeRange.ParagraphFormat.LeftIndent = 18;
        codeRange.ParagraphFormat.RightIndent = 18;
        codeRange.ParagraphFormat.SpaceBefore = 6;
        codeRange.ParagraphFormat.SpaceAfter = 6;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatCodeBlock] Ошибка: {ex.Message}");
    }
}
```

## 🎯 Как это работает (пошагово)

### Пример 1: Блок кода с языком
```
Исходный текст в документе:
```csharp
int x = 10;
```
```

**Шаг 1:** Находим блок кода (позиции 0-20, например)

**Шаг 2:** Применяем форматирование (шрифт Consolas, фон и т.д.)

**Шаг 3:** Удаляем начало:
- Текст начинается с "```csharp\n"
- Находим позицию \n (это позиция 10)
- Удаляем диапазон от 0 до 11 (включая \n)
- Теперь текст: "int x = 10;\n```"

**Шаг 4:** Удаляем конец:
- Текст заканчивается на "\n```"
- Удаляем последние 4 символа
- Теперь текст: "int x = 10;"

**Шаг 5:** Повторно применяем форматирование

### Пример 2: Блок кода без языка
```
Исходный текст:
```
код
```
```

**Шаг 1:** Находим блок кода

**Шаг 2:** Применяем форматирование

**Шаг 3:** Удаляем начало:
- Текст начинается с "```\n"
- Удаляем первые 4 символа
- Текст: "код\n```"

**Шаг 4:** Удаляем конец:
- Удаляем "\n```"
- Текст: "код"

## ⚠️ Важные моменты

1. **После Delete() нужно обновлять codeRange** - позиции сдвигаются!

2. **Проверяем разные варианты окончаний** - может быть `\n````, `\r\n```` или просто ```` 

3. **Порядок важен** - сначала удаляем начало, потом конец

4. **Повторно применяем форматирование** - на случай, если оно сбросилось при удалении

## 🧪 Тестирование

После реализации проверьте:

1. Блок с языком: ` ```csharp\nкод\n``` `
2. Блок без языка: ` ```\nкод\n``` `
3. Однострочный блок: ` ```код``` `
4. Блок с несколькими строками



