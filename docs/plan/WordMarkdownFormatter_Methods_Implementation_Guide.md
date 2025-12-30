# Пошаговая инструкция: Реализация методов FormatHorizontalRule() и RemoveMarkdownSyntax()

## Содержание
1. [Метод FormatHorizontalRule()](#метод-formathorizontalrule)
2. [Метод RemoveMarkdownSyntax()](#метод-removemarkdownsyntax)
3. [Интеграция методов](#интеграция-методов)

---

## Метод FormatHorizontalRule()

### Общее описание
Метод `FormatHorizontalRule()` предназначен для преобразования горизонтальных линий Markdown (`---`, `***`, `___`) в визуальные разделители в документе Word.

### Шаг 1: Добавление метода в класс MarkdownElementFormatter

**Действие:** Откройте файл `Services/WordMarkdownFormatter.cs` и найдите класс `MarkdownElementFormatter`. Добавьте новый метод после метода `FormatTable()` (примерно после строки 904).

**Расположение:** После метода `FormatTable()`, перед закрывающей скобкой класса `MarkdownElementFormatter`

**Код метода:**

```csharp
/// <summary>
/// Применение форматирования для горизонтальной линии
/// </summary>
public void FormatHorizontalRule(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        // Получаем диапазон горизонтальной линии в документе
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range hrRange = _activeDoc.Range(start, end);

        // Получаем параграф, содержащий горизонтальную линию
        Paragraph paragraph = hrRange.Paragraphs[1];

        // Очищаем текст параграфа (удаляем символы ---, *** или ___)
        paragraph.Range.Text = "";

        // Применяем границу снизу параграфа для создания визуальной линии
        paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
        paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth050pt;
        paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorAutomatic;

        // Устанавливаем отступы для визуального выделения
        paragraph.Range.ParagraphFormat.SpaceBefore = 12;  // Отступ сверху
        paragraph.Range.ParagraphFormat.SpaceAfter = 12;   // Отступ снизу
        paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

        // Опционально: можно добавить текст-заполнитель для визуального разделителя
        // paragraph.Range.Text = "─────────────────────────";
        // paragraph.Range.Font.Color = WdColor.wdColorGray50;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatHorizontalRule] Ошибка: {ex.Message}");
    }
}
```

### Шаг 2: Объяснение логики метода

#### 2.1. Получение диапазона
```csharp
int start = documentRange.Start + element.StartPosition;
int end = documentRange.Start + element.EndPosition;
Range hrRange = _activeDoc.Range(start, end);
```
- Вычисляем абсолютные позиции в документе
- Создаем Range для работы с горизонтальной линией

#### 2.2. Получение параграфа
```csharp
Paragraph paragraph = hrRange.Paragraphs[1];
```
- Получаем параграф, содержащий горизонтальную линию
- В Word горизонтальные линии обычно находятся в отдельных параграфах

#### 2.3. Удаление синтаксиса
```csharp
paragraph.Range.Text = "";
```
- Очищаем текст параграфа, удаляя символы `---`, `***` или `___`

#### 2.4. Применение границы
```csharp
paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth050pt;
paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorAutomatic;
```
- Устанавливаем границу снизу параграфа
- Стиль: одинарная линия
- Толщина: 0.5 пункта
- Цвет: автоматический (обычно черный)

#### 2.5. Настройка отступов
```csharp
paragraph.Range.ParagraphFormat.SpaceBefore = 12;  // Отступ сверху
paragraph.Range.ParagraphFormat.SpaceAfter = 12;   // Отступ снизу
paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```
- Добавляем отступы для визуального разделения
- Выравнивание по центру (опционально)

### Шаг 3: Альтернативный вариант (с текстовым разделителем)

Если вы хотите использовать текстовый разделитель вместо границы, используйте этот вариант:

```csharp
/// <summary>
/// Применение форматирования для горизонтальной линии (вариант с текстовым разделителем)
/// </summary>
public void FormatHorizontalRule(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range hrRange = _activeDoc.Range(start, end);

        Paragraph paragraph = hrRange.Paragraphs[1];

        // Заменяем символы на текстовый разделитель
        paragraph.Range.Text = "─────────────────────────";
        
        // Применяем форматирование
        paragraph.Range.Font.Color = WdColor.wdColorGray50;
        paragraph.Range.Font.Size = 8;
        paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        paragraph.Range.ParagraphFormat.SpaceBefore = 12;
        paragraph.Range.ParagraphFormat.SpaceAfter = 12;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatHorizontalRule] Ошибка: {ex.Message}");
    }
}
```

**Рекомендация:** Используйте первый вариант (с границей), так как он более соответствует стандартному форматированию горизонтальных линий в Word.

### Шаг 4: Обновление метода ApplyFormattingToElement()

**Действие:** Найдите метод `ApplyFormattingToElement()` в классе `WordMarkdownFormatter` и обновите обработку `HorizontalRule`.

**Найдите этот код:**
```csharp
case MarkdownElementType.HorizontalRule:
    // Горизонтальные линии можно обработать отдельно
    // Например, вставить разделитель или пустой параграф
    break;
```

**Замените на:**
```csharp
case MarkdownElementType.HorizontalRule:
    _elementFormatter.FormatHorizontalRule(element, documentRange);
    break;
```

---

## Метод RemoveMarkdownSyntax()

### Общее описание
Вспомогательный метод для унификации удаления синтаксических маркеров Markdown из текста документа Word.

### Шаг 1: Добавление метода в класс MarkdownElementFormatter

**Действие:** Добавьте метод `RemoveMarkdownSyntax()` в класс `MarkdownElementFormatter` после метода `LogError()` (примерно после строки 920).

**Расположение:** После метода `LogError()`, перед закрывающей скобкой класса `MarkdownElementFormatter`

**Код метода:**

```csharp
/// <summary>
/// Удаление синтаксических маркеров Markdown из текста
/// </summary>
/// <param name="range">Диапазон текста, из которого нужно удалить синтаксис</param>
/// <param name="syntaxToRemove">Строка синтаксиса для удаления (например, "**", "*", "~~", "`", "```")</param>
/// <returns>true если синтаксис был найден и удален, false в противном случае</returns>
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

### Шаг 2: Улучшенная версия метода (с поддержкой множественных вхождений)

Если нужно удалить синтаксис только в определенных позициях (например, только в начале и конце), используйте эту версию:

```csharp
/// <summary>
/// Удаление синтаксических маркеров Markdown из текста
/// </summary>
/// <param name="range">Диапазон текста, из которого нужно удалить синтаксис</param>
/// <param name="syntaxToRemove">Строка синтаксиса для удаления</param>
/// <param name="removeFromStart">Удалять ли синтаксис только из начала и конца (true) или все вхождения (false)</param>
/// <returns>true если синтаксис был найден и удален, false в противном случае</returns>
public bool RemoveMarkdownSyntax(Range range, string syntaxToRemove, bool removeFromStart = false)
{
    try
    {
        if (range == null || string.IsNullOrEmpty(syntaxToRemove))
            return false;

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
            newText = currentText.Replace(syntaxToRemove, "");
        }

        // Обновляем текст только если он изменился
        if (newText != currentText)
        {
            range.Text = newText;
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

### Шаг 3: Объяснение логики метода

#### 3.1. Валидация параметров
```csharp
if (range == null || string.IsNullOrEmpty(syntaxToRemove))
    return false;
```
- Проверяем, что диапазон и строка синтаксиса не пустые

#### 3.2. Получение текста
```csharp
string currentText = range.Text;
if (string.IsNullOrEmpty(currentText))
    return false;
```
- Получаем текущий текст из диапазона
- Проверяем, что текст не пустой

#### 3.3. Проверка наличия синтаксиса
```csharp
if (!currentText.Contains(syntaxToRemove))
    return false;
```
- Проверяем, содержит ли текст искомый синтаксис
- Если нет - возвращаем false

#### 3.4. Удаление синтаксиса
```csharp
string newText = currentText.Replace(syntaxToRemove, "");
```
- Заменяем все вхождения синтаксиса на пустую строку
- Или используем логику удаления только из начала/конца (в улучшенной версии)

#### 3.5. Обновление диапазона
```csharp
range.Text = newText;
return true;
```
- Обновляем текст в диапазоне
- Возвращаем true, если удаление прошло успешно

### Шаг 4: Примеры использования метода

#### Пример 1: Удаление жирного текста
```csharp
// В методе FormatBoldText()
Range textRange = _activeDoc.Range(start, end);
textRange.Font.Bold = -1;
RemoveMarkdownSyntax(textRange, "**");  // Удаляем **
```

#### Пример 2: Удаление курсива
```csharp
// В методе FormatItalicText()
Range textRange = _activeDoc.Range(start, end);
textRange.Font.Italic = -1;
RemoveMarkdownSyntax(textRange, "*", removeFromStart: true);  // Удаляем только из начала и конца
```

#### Пример 3: Удаление зачеркнутого текста
```csharp
// В методе FormatStrikethroughText()
Range textRange = _activeDoc.Range(start, end);
textRange.Font.StrikeThrough = -1;
RemoveMarkdownSyntax(textRange, "~~");  // Удаляем ~~
```

#### Пример 4: Удаление инлайн-кода
```csharp
// В методе FormatInlineCode()
Range codeRange = _activeDoc.Range(start, end);
codeRange.Font.Name = "Courier New";
RemoveMarkdownSyntax(codeRange, "`", removeFromStart: true);  // Удаляем обратные кавычки
```

### Шаг 5: Опциональное рефакторинг существующих методов

**Примечание:** Метод `RemoveMarkdownSyntax()` можно использовать для рефакторинга существующих методов форматирования, но это не обязательно. Текущая реализация, где каждый метод удаляет синтаксис самостоятельно, также работает корректно.

**Пример рефакторинга метода FormatBoldText():**

**Было:**
```csharp
public void FormatBoldText(MarkdownElementMatch element, Range documentRange)
{
    // ...
    string currentText = textRange.Text;
    string newText = element.Content;
    textRange.Text = newText;
    // ...
}
```

**Можно заменить на:**
```csharp
public void FormatBoldText(MarkdownElementMatch element, Range documentRange)
{
    // ...
    textRange.Font.Bold = -1;
    RemoveMarkdownSyntax(textRange, "**");  // Используем вспомогательный метод
    // ...
}
```

**Рекомендация:** Рефакторинг не обязателен, но может улучшить читаемость кода. Можно оставить текущую реализацию, так как она уже работает.

---

## Интеграция методов

### Шаг 1: Проверка компиляции

**Действие:** После добавления методов проверьте, что проект компилируется без ошибок.

**Команда:** В Visual Studio нажмите `Ctrl+Shift+B` или выберите `Build → Build Solution`

### Шаг 2: Тестирование FormatHorizontalRule()

**Создайте тестовый документ Word с содержимым:**
```
Текст до разделителя

---

Текст после разделителя
```

**Действия:**
1. Откройте тестовый документ в Word
2. Запустите форматирование (через кнопку в Ribbon, когда она будет добавлена)
3. Проверьте, что:
   - Символы `---` удалены
   - Вместо них появилась горизонтальная линия (граница параграфа)
   - Отступы применены корректно

### Шаг 3: Тестирование RemoveMarkdownSyntax()

**Действия:**
1. Создайте тестовый код, который вызывает метод напрямую (для отладки)
2. Проверьте различные варианты:
   - Удаление `**` из текста `**жирный**`
   - Удаление `*` из текста `*курсив*`
   - Удаление `~~` из текста `~~зачеркнутый~~`
   - Удаление `` ` `` из текста `` `код` ``

### Шаг 4: Обработка ошибок

**Убедитесь, что:**
- Все исключения обрабатываются в блоках `try-catch`
- Ошибки логируются через `Debug.WriteLine()` или `LogError()`
- Методы не падают при некорректных входных данных

---

## Полный код методов

### FormatHorizontalRule() - финальная версия

```csharp
/// <summary>
/// Применение форматирования для горизонтальной линии
/// </summary>
public void FormatHorizontalRule(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        // Получаем диапазон горизонтальной линии в документе
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range hrRange = _activeDoc.Range(start, end);

        // Получаем параграф, содержащий горизонтальную линию
        Paragraph paragraph = hrRange.Paragraphs[1];

        // Очищаем текст параграфа (удаляем символы ---, *** или ___)
        paragraph.Range.Text = "";

        // Применяем границу снизу параграфа для создания визуальной линии
        paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
        paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth050pt;
        paragraph.Range.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorAutomatic;

        // Устанавливаем отступы для визуального выделения
        paragraph.Range.ParagraphFormat.SpaceBefore = 12;  // Отступ сверху (1 пункт)
        paragraph.Range.ParagraphFormat.SpaceAfter = 12;   // Отступ снизу (1 пункт)
        paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatHorizontalRule] Ошибка: {ex.Message}");
    }
}
```

### RemoveMarkdownSyntax() - финальная версия (упрощенная)

```csharp
/// <summary>
/// Удаление синтаксических маркеров Markdown из текста
/// </summary>
/// <param name="range">Диапазон текста, из которого нужно удалить синтаксис</param>
/// <param name="syntaxToRemove">Строка синтаксиса для удаления (например, "**", "*", "~~", "`")</param>
/// <returns>true если синтаксис был найден и удален, false в противном случае</returns>
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

### RemoveMarkdownSyntax() - улучшенная версия (с опцией удаления только из начала/конца)

```csharp
/// <summary>
/// Удаление синтаксических маркеров Markdown из текста
/// </summary>
/// <param name="range">Диапазон текста, из которого нужно удалить синтаксис</param>
/// <param name="syntaxToRemove">Строка синтаксиса для удаления</param>
/// <param name="removeFromStart">Удалять ли синтаксис только из начала и конца (true) или все вхождения (false)</param>
/// <returns>true если синтаксис был найден и удален, false в противном случае</returns>
public bool RemoveMarkdownSyntax(Range range, string syntaxToRemove, bool removeFromStart = false)
{
    try
    {
        if (range == null || string.IsNullOrEmpty(syntaxToRemove))
            return false;

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
            newText = currentText.Replace(syntaxToRemove, "");
        }

        // Обновляем текст только если он изменился
        if (newText != currentText)
        {
            range.Text = newText;
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

---

## Заключение

После реализации этих методов:

1. ✅ **FormatHorizontalRule()** - горизонтальные линии будут корректно форматироваться в Word
2. ✅ **RemoveMarkdownSyntax()** - появится унифицированный способ удаления синтаксиса (опционально для рефакторинга)

**Следующие шаги:**
- Интеграция с Ribbon (добавление кнопки)
- Улучшение реализации таблиц
- Тестирование всех методов

---

**Дата создания инструкции:** 2024
**Версия:** 1.0

