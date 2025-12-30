# Пошаговая инструкция: Реализация удаления синтаксиса Markdown

## 📋 Обзор

Эта инструкция описывает пошаговую реализацию исправления проблемы с удалением синтаксических маркеров Markdown после форматирования в документе Word.

**Цель:** После форматирования в документе должны остаться только форматированные элементы Word без синтаксических маркеров Markdown (`#`, `*`, `**`, `~~`, `` ` ``, и т.д.).

**Подход:** Использование улучшенного метода `RemoveMarkdownSyntax` и правильная обработка позиций элементов.

---

## 🎯 Этап 1: Улучшение метода RemoveMarkdownSyntax

### Шаг 1.1: Открыть файл WordMarkdownFormatter.cs

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 949-979

### Шаг 1.2: Заменить существующий метод RemoveMarkdownSyntax

**Текущий код:**
```csharp
public bool RemoveMarkdownSyntax(Range range, string syntaxToRemove)
{
    try
    {
        if (range == null || string.IsNullOrEmpty(syntaxToRemove))
            return false;

        // Получаем текущий текст диапазона
        string currentText = range.Text;

        if (string.IsNullOrEmpty(currentText))
            return false;

        // Проверяем, содержит ли текст синтаксис для удаления
        if (!currentText.Contains(syntaxToRemove))
            return false;

        // Удаляем все вхождения синтаксиса
        string newText = currentText.Replace(syntaxToRemove, "");

        // Обновляем текст в диапазоне
        range.Text = newText;

        return true;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[RemoveMarkdownSyntax] Ошибка при удалении синтаксиса '{syntaxToRemove}': {ex.Message}");
        return false;
    }
}
```

**Новый код:**
```csharp
/// <summary>
/// Удаление синтаксических маркеров Markdown из текста
/// </summary>
/// <param name="range">Диапазон текста, из которого нужно удалить синтаксис</param>
/// <param name="syntaxToRemove">Строка синтаксиса для удаления (например, "**", "*", "~~", "`")</param>
/// <param name="removeFromStart">Удалять ли синтаксис только из начала и конца (true) или все вхождения (false)</param>
/// <returns>true если синтаксис был найден и удален, false в противном случае</returns>
public bool RemoveMarkdownSyntax(Range range, string syntaxToRemove, bool removeFromStart = false)
{
    try
    {
        if (range == null || string.IsNullOrEmpty(syntaxToRemove))
            return false;

        // Получаем текущий текст диапазона
        string currentText = range.Text;

        if (string.IsNullOrEmpty(currentText))
            return false;

        string newText;

        if (removeFromStart)
        {
            // Удаляем синтаксис только из начала и конца
            newText = currentText.Trim();
            
            if (newText.StartsWith(syntaxToRemove))
            {
                newText = newText.Substring(syntaxToRemove.Length);
            }
            
            if (newText.EndsWith(syntaxToRemove))
            {
                newText = newText.Substring(0, newText.Length - syntaxToRemove.Length);
            }
        }
        else
        {
            // Удаляем все вхождения синтаксиса
            if (!currentText.Contains(syntaxToRemove))
                return false;
            
            newText = currentText.Replace(syntaxToRemove, "");
        }

        // Обновляем текст только если он изменился
        if (newText != currentText)
        {
            range.Text = newText;
            Debug.WriteLine($"[RemoveMarkdownSyntax] Удален синтаксис '{syntaxToRemove}' из диапазона. Длина изменена: {currentText.Length} -> {newText.Length}");
            return true;
        }

        return false;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[RemoveMarkdownSyntax] Ошибка при удалении синтаксиса '{syntaxToRemove}': {ex.Message}");
        return false;
    }
}
```

### Шаг 1.3: Проверка компиляции

1. Открыть проект в Visual Studio
2. Нажать `Ctrl+Shift+B` для компиляции
3. Убедиться, что нет ошибок компиляции
4. Если есть ошибки, проверить правильность замены кода

### Шаг 1.4: Тестирование метода RemoveMarkdownSyntax

**Создать тестовый метод (временно, для проверки):**

```csharp
// Временный метод для тестирования - удалить после проверки
public void TestRemoveMarkdownSyntax()
{
    Range testRange = _activeDoc.Range(0, 10);
    testRange.Text = "**жирный**";
    
    bool result = RemoveMarkdownSyntax(testRange, "**", true);
    Debug.WriteLine($"Результат: {result}, Текст: '{testRange.Text}'");
    // Ожидаемый результат: result = true, Текст = "жирный"
}
```

---

## 🎯 Этап 2: Исправление метода FormatHeading

### Шаг 2.1: Найти метод FormatHeading

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 527-573

### Шаг 2.2: Заменить метод FormatHeading

**Текущий код:**
```csharp
public void FormatHeading(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null || !element.Metadata.ContainsKey("Level"))
            return;

        int level = (int)element.Metadata["Level"];
        if (level < 1 || level > 6)
            return;

        // Получаем диапазон заголовка в документе
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range headingRange = _activeDoc.Range(start, end);

        // Определяем стиль заголовка
        WdBuiltinStyle headingStyle;
        switch (level)
        {
            case 1: headingStyle = WdBuiltinStyle.wdStyleHeading1; break;
            case 2: headingStyle = WdBuiltinStyle.wdStyleHeading2; break;
            case 3: headingStyle = WdBuiltinStyle.wdStyleHeading3; break;
            case 4: headingStyle = WdBuiltinStyle.wdStyleHeading4; break;
            case 5: headingStyle = WdBuiltinStyle.wdStyleHeading5; break;
            case 6: headingStyle = WdBuiltinStyle.wdStyleHeading6; break;
            default: headingStyle = WdBuiltinStyle.wdStyleNormal; break;
        }

        // Применяем стиль
        headingRange.set_Style(headingStyle);

        // Удаляем символы # и пробелы из начала
        string currentText = headingRange.Text;
        string newText = element.Content;

        // Заменяем текст, сохраняя форматирование
        headingRange.Text = newText;

        // Повторно применяем стиль после замены текста
        headingRange.set_Style(headingStyle);
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatHeading] Ошибка: {ex.Message}");
    }
}
```

**Новый код:**
```csharp
/// <summary>
/// Применение стиля заголовка к найденному элементу
/// </summary>
public void FormatHeading(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null || !element.Metadata.ContainsKey("Level"))
            return;

        int level = (int)element.Metadata["Level"];
        if (level < 1 || level > 6)
            return;

        // Получаем диапазон заголовка в документе
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range headingRange = _activeDoc.Range(start, end);

        // Определяем стиль заголовка
        WdBuiltinStyle headingStyle;
        switch (level)
        {
            case 1: headingStyle = WdBuiltinStyle.wdStyleHeading1; break;
            case 2: headingStyle = WdBuiltinStyle.wdStyleHeading2; break;
            case 3: headingStyle = WdBuiltinStyle.wdStyleHeading3; break;
            case 4: headingStyle = WdBuiltinStyle.wdStyleHeading4; break;
            case 5: headingStyle = WdBuiltinStyle.wdStyleHeading5; break;
            case 6: headingStyle = WdBuiltinStyle.wdStyleHeading6; break;
            default: headingStyle = WdBuiltinStyle.wdStyleNormal; break;
        }

        // Применяем стиль
        headingRange.set_Style(headingStyle);

        // Удаляем символы # и пробелы из начала заголовка
        string headingText = headingRange.Text;
        if (!string.IsNullOrEmpty(headingText))
        {
            // Находим количество символов # в начале
            int hashCount = 0;
            while (hashCount < headingText.Length && headingText[hashCount] == '#')
            {
                hashCount++;
            }

            if (hashCount > 0)
            {
                // Вычисляем длину удаляемой части (# и пробел после них)
                int removeLength = hashCount;
                if (removeLength < headingText.Length && headingText[removeLength] == ' ')
                {
                    removeLength++;
                }

                // Создаем диапазон для удаления символов #
                Range removeRange = _activeDoc.Range(
                    headingRange.Start,
                    headingRange.Start + removeLength
                );
                removeRange.Delete();

                Debug.WriteLine($"[FormatHeading] Удалено {removeLength} символов из заголовка уровня {level}");
            }
        }

        // Повторно применяем стиль после удаления символов
        headingRange.set_Style(headingStyle);
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatHeading] Ошибка: {ex.Message}");
    }
}
```

### Шаг 2.3: Проверка компиляции

1. Скомпилировать проект
2. Убедиться, что нет ошибок

### Шаг 2.4: Тестирование FormatHeading

**Тестовые случаи:**
1. `# Заголовок 1` → должно стать "Заголовок 1" со стилем Heading1
2. `## Заголовок 2` → должно стать "Заголовок 2" со стилем Heading2
3. `### Заголовок 3` → должно стать "Заголовок 3" со стилем Heading3

---

## 🎯 Этап 3: Исправление метода FormatBoldText

### Шаг 3.1: Найти метод FormatBoldText

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 578-604

### Шаг 3.2: Заменить метод FormatBoldText

**Текущий код:**
```csharp
public void FormatBoldText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем жирный шрифт
        textRange.Font.Bold = -1;

        // Удаляем символы **
        string currentText = textRange.Text;
        string newText = element.Content;
        textRange.Text = newText;

        // Повторно применяем форматирование
        textRange.Font.Bold = -1;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatBoldText] Ошибка: {ex.Message}");
    }
}
```

**Новый код:**
```csharp
/// <summary>
/// Применение жирного форматирования
/// </summary>
public void FormatBoldText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем жирный шрифт
        textRange.Font.Bold = -1;

        // Удаляем символы ** из начала и конца
        // Используем улучшенный метод RemoveMarkdownSyntax
        bool removedStart = RemoveMarkdownSyntax(
            _activeDoc.Range(textRange.Start, textRange.Start + 2),
            "**",
            true
        );
        
        if (!removedStart)
        {
            // Если не удалось удалить из начала, пробуем удалить все вхождения
            RemoveMarkdownSyntax(textRange, "**", false);
        }
        else
        {
            // Обновляем диапазон после удаления начала
            textRange = _activeDoc.Range(textRange.Start, textRange.End - 2);
            
            // Удаляем ** из конца
            if (textRange.Text.EndsWith("**"))
            {
                Range endRange = _activeDoc.Range(textRange.End - 2, textRange.End);
                endRange.Delete();
            }
        }

        // Повторно применяем форматирование
        textRange.Font.Bold = -1;
        
        Debug.WriteLine($"[FormatBoldText] Применено жирное форматирование, удалены маркеры **");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatBoldText] Ошибка: {ex.Message}");
    }
}
```

**Альтернативный вариант (более простой и надежный):**
```csharp
/// <summary>
/// Применение жирного форматирования
/// </summary>
public void FormatBoldText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем жирный шрифт
        textRange.Font.Bold = -1;

        // Удаляем символы ** из текста
        string currentText = textRange.Text;
        if (!string.IsNullOrEmpty(currentText))
        {
            // Удаляем ** из начала
            if (currentText.StartsWith("**"))
            {
                Range startRange = _activeDoc.Range(textRange.Start, textRange.Start + 2);
                startRange.Delete();
                // Обновляем диапазон
                textRange = _activeDoc.Range(textRange.Start, textRange.End - 2);
            }

            // Удаляем ** из конца
            string updatedText = textRange.Text;
            if (updatedText.EndsWith("**"))
            {
                Range endRange = _activeDoc.Range(textRange.End - 2, textRange.End);
                endRange.Delete();
            }
        }

        // Повторно применяем форматирование
        textRange.Font.Bold = -1;
        
        Debug.WriteLine($"[FormatBoldText] Применено жирное форматирование, удалены маркеры **");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatBoldText] Ошибка: {ex.Message}");
    }
}
```

### Шаг 3.3: Проверка компиляции

1. Скомпилировать проект
2. Убедиться, что нет ошибок

### Шаг 3.4: Тестирование FormatBoldText

**Тестовые случаи:**
1. `**жирный текст**` → должно стать "жирный текст" с жирным форматированием
2. `**жирный**` → должно стать "жирный" с жирным форматированием

---

## 🎯 Этап 4: Исправление метода FormatItalicText

### Шаг 4.1: Найти метод FormatItalicText

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 609-634

### Шаг 4.2: Заменить метод FormatItalicText

**Текущий код:**
```csharp
public void FormatItalicText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем курсив
        textRange.Font.Italic = -1;

        // Удаляем символы *
        string newText = element.Content;
        textRange.Text = newText;

        // Повторно применяем форматирование
        textRange.Font.Italic = -1;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatItalicText] Ошибка: {ex.Message}");
    }
}
```

**Новый код:**
```csharp
/// <summary>
/// Применение курсива
/// </summary>
public void FormatItalicText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем курсив
        textRange.Font.Italic = -1;

        // Удаляем символы * из текста
        string currentText = textRange.Text;
        if (!string.IsNullOrEmpty(currentText))
        {
            // Удаляем * из начала (но не **)
            if (currentText.StartsWith("*") && !currentText.StartsWith("**"))
            {
                Range startRange = _activeDoc.Range(textRange.Start, textRange.Start + 1);
                startRange.Delete();
                // Обновляем диапазон
                textRange = _activeDoc.Range(textRange.Start, textRange.End - 1);
            }

            // Удаляем * из конца (но не **)
            string updatedText = textRange.Text;
            if (updatedText.EndsWith("*") && !updatedText.EndsWith("**"))
            {
                Range endRange = _activeDoc.Range(textRange.End - 1, textRange.End);
                endRange.Delete();
            }
        }

        // Повторно применяем форматирование
        textRange.Font.Italic = -1;
        
        Debug.WriteLine($"[FormatItalicText] Применен курсив, удалены маркеры *");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatItalicText] Ошибка: {ex.Message}");
    }
}
```

### Шаг 4.3: Проверка компиляции

1. Скомпилировать проект
2. Убедиться, что нет ошибок

### Шаг 4.4: Тестирование FormatItalicText

**Тестовые случаи:**
1. `*курсив*` → должно стать "курсив" с курсивом
2. `*курсивный текст*` → должно стать "курсивный текст" с курсивом

---

## 🎯 Этап 5: Исправление метода FormatStrikethroughText

### Шаг 5.1: Найти метод FormatStrikethroughText

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 639-664

### Шаг 5.2: Заменить метод FormatStrikethroughText

**Текущий код:**
```csharp
public void FormatStrikethroughText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем зачеркивание
        textRange.Font.StrikeThrough = -1;

        // Удаляем символы ~~
        string newText = element.Content;
        textRange.Text = newText;

        // Повторно применяем форматирование
        textRange.Font.StrikeThrough = -1;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatStrikethroughText] Ошибка: {ex.Message}");
    }
}
```

**Новый код:**
```csharp
/// <summary>
/// Применение зачеркивания
/// </summary>
public void FormatStrikethroughText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем зачеркивание
        textRange.Font.StrikeThrough = -1;

        // Удаляем символы ~~ из текста
        string currentText = textRange.Text;
        if (!string.IsNullOrEmpty(currentText))
        {
            // Удаляем ~~ из начала
            if (currentText.StartsWith("~~"))
            {
                Range startRange = _activeDoc.Range(textRange.Start, textRange.Start + 2);
                startRange.Delete();
                // Обновляем диапазон
                textRange = _activeDoc.Range(textRange.Start, textRange.End - 2);
            }

            // Удаляем ~~ из конца
            string updatedText = textRange.Text;
            if (updatedText.EndsWith("~~"))
            {
                Range endRange = _activeDoc.Range(textRange.End - 2, textRange.End);
                endRange.Delete();
            }
        }

        // Повторно применяем форматирование
        textRange.Font.StrikeThrough = -1;
        
        Debug.WriteLine($"[FormatStrikethroughText] Применено зачеркивание, удалены маркеры ~~");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatStrikethroughText] Ошибка: {ex.Message}");
    }
}
```

### Шаг 5.3: Проверка компиляции

1. Скомпилировать проект
2. Убедиться, что нет ошибок

---

## 🎯 Этап 6: Исправление метода FormatInlineCode

### Шаг 6.1: Найти метод FormatInlineCode

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 669-700

### Шаг 6.2: Заменить метод FormatInlineCode

**Текущий код:**
```csharp
public void FormatInlineCode(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range codeRange = _activeDoc.Range(start, end);

        // Устанавливаем моноширинный шрифт
        codeRange.Font.Name = "Courier New";
        codeRange.Font.Size = 10;

        // Опционально: добавляем фон
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;

        // Удаляем обратные кавычки
        string newText = element.Content;
        codeRange.Text = newText;

        // Повторно применяем форматирование
        codeRange.Font.Name = "Courier New";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatInlineCode] Ошибка: {ex.Message}");
    }
}
```

**Новый код:**
```csharp
/// <summary>
/// Применение форматирования для инлайн-кода
/// </summary>
public void FormatInlineCode(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range codeRange = _activeDoc.Range(start, end);

        // Устанавливаем моноширинный шрифт
        codeRange.Font.Name = "Courier New";
        codeRange.Font.Size = 10;

        // Опционально: добавляем фон
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;

        // Удаляем обратные кавычки ` из текста
        string currentText = codeRange.Text;
        if (!string.IsNullOrEmpty(currentText))
        {
            // Удаляем ` из начала
            if (currentText.StartsWith("`"))
            {
                Range startRange = _activeDoc.Range(codeRange.Start, codeRange.Start + 1);
                startRange.Delete();
                // Обновляем диапазон
                codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - 1);
            }

            // Удаляем ` из конца
            string updatedText = codeRange.Text;
            if (updatedText.EndsWith("`"))
            {
                Range endRange = _activeDoc.Range(codeRange.End - 1, codeRange.End);
                endRange.Delete();
            }
        }

        // Повторно применяем форматирование
        codeRange.Font.Name = "Courier New";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        
        Debug.WriteLine($"[FormatInlineCode] Применено форматирование кода, удалены маркеры `");
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatInlineCode] Ошибка: {ex.Message}");
    }
}
```

### Шаг 6.3: Проверка компиляции

1. Скомпилировать проект
2. Убедиться, что нет ошибок

---

## 🎯 Этап 7: Исправление метода FormatCodeBlock

### Шаг 7.1: Найти метод FormatCodeBlock

**Файл:** `Services/WordMarkdownFormatter.cs`  
**Строки:** 705-750 (примерно)

### Шаг 7.2: Изучить текущую реализацию

Прочитать метод `FormatCodeBlock` и понять его структуру.

### Шаг 7.3: Добавить удаление маркеров ``` 

**Добавить после применения форматирования, перед заменой текста:**

```csharp
// Удаляем маркеры ``` из начала и конца блока кода
string currentText = codeRange.Text;
if (!string.IsNullOrEmpty(currentText))
{
    // Удаляем ``` из начала (может быть с языком: ```csharp)
    if (currentText.StartsWith("```"))
    {
        // Находим конец первой строки (до \n или \r\n)
        int firstLineEnd = currentText.IndexOfAny(new[] { '\n', '\r' });
        if (firstLineEnd > 0)
        {
            Range startRange = _activeDoc.Range(codeRange.Start, codeRange.Start + firstLineEnd + 1);
            startRange.Delete();
            // Обновляем диапазон
            codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - (firstLineEnd + 1));
        }
        else
        {
            // Если нет перевода строки, удаляем только ```
            Range startRange = _activeDoc.Range(codeRange.Start, codeRange.Start + 3);
            startRange.Delete();
            codeRange = _activeDoc.Range(codeRange.Start, codeRange.End - 3);
        }
    }

    // Удаляем ``` из конца
    string updatedText = codeRange.Text;
    if (updatedText.EndsWith("```"))
    {
        Range endRange = _activeDoc.Range(codeRange.End - 3, codeRange.End);
        endRange.Delete();
    }
    else if (updatedText.EndsWith("\n```") || updatedText.EndsWith("\r\n```"))
    {
        // Удаляем перевод строки и ```
        int removeLength = updatedText.EndsWith("\r\n```") ? 5 : 4;
        Range endRange = _activeDoc.Range(codeRange.End - removeLength, codeRange.End);
        endRange.Delete();
    }
}
```

---

## 🎯 Этап 8: Исправление остальных методов форматирования

### Шаг 8.1: Найти методы FormatLink, FormatListItem, FormatQuote, FormatTable, FormatHorizontalRule

**Файл:** `Services/WordMarkdownFormatter.cs`

### Шаг 8.2: Применить аналогичные исправления

Для каждого метода:
1. Найти место, где заменяется текст
2. Добавить удаление соответствующих маркеров синтаксиса
3. Использовать подход удаления из начала и конца диапазона

**Пример для FormatLink:**
```csharp
// Удаляем [ и ] из начала и конца
if (currentText.StartsWith("["))
{
    Range startRange = _activeDoc.Range(linkRange.Start, linkRange.Start + 1);
    startRange.Delete();
    linkRange = _activeDoc.Range(linkRange.Start, linkRange.End - 1);
}

// Удаляем (url) часть
// ... логика удаления URL части ...
```

---

## 🎯 Этап 9: Комплексное тестирование

### Шаг 9.1: Создать тестовый документ Word

Создать документ со следующим содержимым:
```
# Заголовок 1

## Заголовок 2

Это **жирный** текст и *курсивный* текст.

Также есть ~~зачеркнутый~~ текст и `код`.

```csharp
int x = 10;
```

- Элемент списка с **жирным** текстом
- Еще один элемент
```

### Шаг 9.2: Применить форматирование

1. Выделить весь текст
2. Нажать кнопку "Форматировать Markdown"
3. Проверить результат

### Шаг 9.3: Проверить результаты

**Ожидаемые результаты:**
- `# Заголовок 1` → "Заголовок 1" (без #, со стилем Heading1)
- `**жирный**` → "жирный" (без **, с жирным форматированием)
- `*курсивный*` → "курсивный" (без *, с курсивом)
- `~~зачеркнутый~~` → "зачеркнутый" (без ~~, с зачеркиванием)
- `` `код` `` → "код" (без `, с моноширинным шрифтом)
- Блок кода без ``` в начале и конце

### Шаг 9.4: Исправить найденные проблемы

Если найдены проблемы:
1. Записать описание проблемы
2. Найти соответствующий метод форматирования
3. Исправить логику удаления маркеров
4. Повторить тестирование

---

## 🎯 Этап 10: Оптимизация и улучшения

### Шаг 10.1: Добавить обработку вложенных элементов

**Проблема:** Если заголовок содержит жирный текст, маркеры жирного текста могут остаться.

**Решение:** Обрабатывать элементы в правильном порядке (сначала родительские, потом вложенные).

### Шаг 10.2: Улучшить обработку ошибок

Добавить более детальное логирование:
```csharp
catch (Exception ex)
{
    Debug.WriteLine($"[FormatBoldText] Ошибка при обработке элемента: {ex.Message}");
    Debug.WriteLine($"[FormatBoldText] StartPosition: {element.StartPosition}, EndPosition: {element.EndPosition}");
    Debug.WriteLine($"[FormatBoldText] Content: '{element.Content}', FullMatch: '{element.FullMatch}'");
    Debug.WriteLine($"[FormatBoldText] StackTrace: {ex.StackTrace}");
}
```

### Шаг 10.3: Добавить проверки граничных случаев

- Пустой текст
- Текст без маркеров
- Неправильный синтаксис
- Очень длинный текст

---

## 📋 Чек-лист выполнения

### Этап 1: RemoveMarkdownSyntax
- [ ] Заменен метод RemoveMarkdownSyntax
- [ ] Добавлен параметр removeFromStart
- [ ] Добавлено логирование
- [ ] Протестирован метод

### Этап 2: FormatHeading
- [ ] Заменен метод FormatHeading
- [ ] Добавлено удаление символов #
- [ ] Протестирован метод

### Этап 3: FormatBoldText
- [ ] Заменен метод FormatBoldText
- [ ] Добавлено удаление символов **
- [ ] Протестирован метод

### Этап 4: FormatItalicText
- [ ] Заменен метод FormatItalicText
- [ ] Добавлено удаление символов *
- [ ] Протестирован метод

### Этап 5: FormatStrikethroughText
- [ ] Заменен метод FormatStrikethroughText
- [ ] Добавлено удаление символов ~~
- [ ] Протестирован метод

### Этап 6: FormatInlineCode
- [ ] Заменен метод FormatInlineCode
- [ ] Добавлено удаление символов `
- [ ] Протестирован метод

### Этап 7: FormatCodeBlock
- [ ] Заменен метод FormatCodeBlock
- [ ] Добавлено удаление символов ```
- [ ] Протестирован метод

### Этап 8: Остальные методы
- [ ] Исправлен FormatLink
- [ ] Исправлен FormatListItem
- [ ] Исправлен FormatQuote
- [ ] Исправлен FormatTable
- [ ] Исправлен FormatHorizontalRule

### Этап 9: Тестирование
- [ ] Создан тестовый документ
- [ ] Протестированы простые элементы
- [ ] Протестированы вложенные элементы
- [ ] Протестированы сложные элементы
- [ ] Протестированы граничные случаи

### Этап 10: Оптимизация
- [ ] Добавлена обработка вложенных элементов
- [ ] Улучшена обработка ошибок
- [ ] Добавлены проверки граничных случаев

---

## 🐛 Решение типичных проблем

### Проблема 1: Позиции элементов сдвигаются после удаления

**Симптом:** Элементы форматируются неправильно после первого удаления.

**Решение:** Убедиться, что элементы обрабатываются в обратном порядке (с конца к началу). Это уже реализовано в методе `FormatMarkdownInWord`.

### Проблема 2: Маркеры удаляются, но форматирование теряется

**Симптом:** После удаления маркеров форматирование Word не применяется.

**Решение:** Убедиться, что форматирование применяется ДО удаления маркеров, и повторно применяется ПОСЛЕ удаления.

### Проблема 3: Ошибка "Range не найден"

**Симптом:** Исключение при попытке создать Range.

**Решение:** Проверить, что позиции вычисляются правильно и не выходят за границы документа.

### Проблема 4: Маркеры удаляются частично

**Симптом:** Удаляется только один маркер из пары (например, только один `*` вместо `**`).

**Решение:** Убедиться, что удаляются оба маркера (начало и конец) отдельными операциями.

---

## 📝 Примечания

1. **Порядок обработки:** Элементы обрабатываются в обратном порядке (с конца к началу), чтобы избежать сдвига позиций.

2. **Логирование:** Все методы форматирования должны логировать свои действия для отладки.

3. **Обработка ошибок:** Каждый метод должен иметь try-catch блок для обработки ошибок.

4. **Тестирование:** После каждого изменения необходимо тестировать на реальных документах.

5. **Производительность:** Для больших документов может потребоваться оптимизация.

---

## ✅ Критерии готовности

Реализация считается завершенной, когда:

1. ✅ Все методы форматирования исправлены
2. ✅ Все маркеры синтаксиса удаляются корректно
3. ✅ Форматирование Word применяется правильно
4. ✅ Все тесты проходят успешно
5. ✅ Нет ошибок компиляции
6. ✅ Нет критических ошибок при выполнении
7. ✅ Логирование работает корректно

---

**Дата создания:** 2024  
**Автор:** AI Assistant  
**Версия:** 1.0

