# Как работает метод ApplyToWord() класса WordTitle

## Введение

Метод `ApplyToWord()` класса `WordTitle` отвечает за преобразование Markdown заголовков в форматированные заголовки в документе Microsoft Word. В этом документе мы разберем, как работает этот метод на примере преобразования следующего Markdown текста:

```markdown
# Модуль DocumentSyncService

## Описание
Сервис для синхронизации Markdown контента с Word документом через CustomXMLPart. Обеспечивает сохранение и загрузку Markdown данных в/из Word документов, что позволяет сохранять исходный Markdown вместе с документом.
```

---

## Архитектура размещения данных в Word

### Концепция Range и Paragraph

В Microsoft Word Interop API данные размещаются через концепцию **Range** (диапазон) и **Paragraph** (параграф):

- **Range** - это непрерывная область текста в документе, определяемая начальной и конечной позицией (в символах)
- **Paragraph** - это структурная единица документа, которая содержит один или несколько Range
- Каждый Range имеет свойства `Start` и `End`, которые указывают позиции символов в документе

### Структура документа Word

```
Документ Word
├── Content (главный Range всего документа)
│   ├── Paragraph 1 (Range: 0-50)
│   │   └── Text: "# Модуль DocumentSyncService"
│   ├── Paragraph 2 (Range: 51-100)
│   │   └── Text: "" (пустой параграф после заголовка)
│   ├── Paragraph 3 (Range: 101-150)
│   │   └── Text: "## Описание"
│   └── ...
```

---

## Пошаговое выполнение метода ApplyToWord()

### Шаг 1: Создание параграфа

```csharp
var paragraph = doc.Content.Paragraphs.Add();
paragraph.Range.ListFormat.RemoveNumbers();
```

**Что происходит:**
- Создается новый параграф в конце документа через `doc.Content.Paragraphs.Add()`
- `doc.Content` - это главный Range всего документа
- Новый параграф получает свой собственный Range, который изначально пустой
- `RemoveNumbers()` убирает возможное наследование нумерации от предыдущих элементов

**На примере нашего текста:**
- Для `# Модуль DocumentSyncService` создается Paragraph 1
- Range этого параграфа изначально пустой (например, Start=0, End=0)

---

### Шаг 2: Вставка текста заголовка

Метод проверяет, есть ли форматированный текст (`Content`) или простой текст (`Text`):

#### Вариант A: Форматированный текст (Content)

```csharp
if (Content != null && Content.Runs.Count > 0)
{
    Content.ApplyToWord(doc, paragraph.Range);
    textRange = paragraph.Range;
}
```

**Что происходит:**
- Если есть `Content` (объект `WordFormattedText`), вызывается его метод `ApplyToWord()`
- `WordFormattedText` содержит коллекцию `Runs` - фрагментов текста с разным форматированием
- Каждый Run вставляется последовательно в `paragraph.Range`

**Как работает WordFormattedText.ApplyToWord():**

1. **Запоминание позиции:**
   ```csharp
   int start = range.End;  // Позиция начала вставки
   ```

2. **Вставка текста:**
   ```csharp
   range.InsertAfter(run.Text);  // Вставляет текст после текущей позиции
   int end = start + run.Text.Length;  // Позиция конца вставленного текста
   ```

3. **Создание Range для вставленного текста:**
   ```csharp
   var insertedRange = doc.Range(start, end);
   ```

4. **Применение форматирования:**
   ```csharp
   insertedRange.Font.Bold = run.IsBold ? -1 : 0;
   insertedRange.Font.Italic = run.IsItalic ? -1 : 0;
   // и т.д.
   ```

5. **Сдвиг Range для следующего Run:**
   ```csharp
   range.SetRange(end, end);  // Перемещаем курсор в конец вставленного текста
   ```

**На примере нашего текста:**
- Для `# Модуль DocumentSyncService`:
  - Если текст простой (без форматирования), создается один Run с текстом "Модуль DocumentSyncService"
  - Run вставляется в paragraph.Range
  - После вставки paragraph.Range содержит текст (например, Start=0, End=25)

#### Вариант B: Простой текст (Text)

```csharp
else if (!string.IsNullOrEmpty(Text))
{
    paragraph.Range.Text = Text;
    textRange = paragraph.Range;
}
```

**Что происходит:**
- Если `Content` отсутствует, используется простое свойство `Text`
- Текст присваивается напрямую через `paragraph.Range.Text`
- Это самый простой способ вставки текста без форматирования

**На примере нашего текста:**
- Для `# Модуль DocumentSyncService`:
  - `Text = "Модуль DocumentSyncService"`
  - `paragraph.Range.Text = "Модуль DocumentSyncService"`
  - Range автоматически расширяется (Start=0, End=25)

---

### Шаг 3: Применение форматирования заголовка

После вставки текста применяется форматирование заголовка:

#### 3.1. Установка размера шрифта

```csharp
float fontSize = GetFontSizeForLevel(Level);
textRange.Font.Size = fontSize;
```

**Что происходит:**
- Метод `GetFontSizeForLevel()` возвращает размер шрифта в зависимости от уровня:
  - Level 1 → 24pt
  - Level 2 → 18pt
  - Level 3 → 14pt
  - и т.д.

**На примере нашего текста:**
- `# Модуль DocumentSyncService` (Level=1) → Font.Size = 24
- `## Описание` (Level=2) → Font.Size = 18

#### 3.2. Применение жирного начертания

```csharp
textRange.Font.Bold = -1;
```

**Что происходит:**
- В Word Interop `-1` означает включение свойства (True)
- `0` означает выключение (False)
- Все символы в `textRange` становятся жирными

#### 3.3. Применение стиля заголовка

```csharp
WdBuiltinStyle headingStyle;
switch (Level)
{
    case 1: headingStyle = WdBuiltinStyle.wdStyleHeading1; break;
    case 2: headingStyle = WdBuiltinStyle.wdStyleHeading2; break;
    // ...
}
textRange.set_Style(headingStyle);
```

**Что происходит:**
- Применяется встроенный стиль Word для заголовка соответствующего уровня
- Стили Word содержат предопределенное форматирование (шрифт, размер, отступы, интервалы)
- Это важно для автоматического создания оглавления в Word

**На примере нашего текста:**
- `# Модуль DocumentSyncService` → `wdStyleHeading1`
- `## Описание` → `wdStyleHeading2`

#### 3.4. Очистка нумерации

```csharp
textRange.ListFormat.RemoveNumbers();
```

**Что происходит:**
- Убирается возможная нумерация, которая могла быть унаследована от стиля
- Это гарантирует, что заголовок не будет частью списка

---

### Шаг 4: Добавление переноса строки

```csharp
paragraph.Range.InsertParagraphAfter();
```

**Что происходит:**
- После заголовка создается новый пустой параграф
- Это обеспечивает визуальное разделение между заголовком и следующим содержимым
- В Word каждый параграф заканчивается маркером параграфа (¶)

**На примере нашего текста:**
- После `# Модуль DocumentSyncService` создается пустой Paragraph 2
- После `## Описание` создается пустой Paragraph 4

#### Очистка нумерации у нового параграфа

```csharp
var lastParagraphIndex = doc.Content.Paragraphs.Count;
var newParagraph = doc.Content.Paragraphs[lastParagraphIndex];
newParagraph.Range.ListFormat.RemoveNumbers();
```

**Что происходит:**
- Новый параграф, созданный через `InsertParagraphAfter()`, также очищается от нумерации
- Это предотвращает наследование форматирования списка следующими элементами

---

## Визуализация процесса на примере

### Преобразование "# Модуль DocumentSyncService"

**Исходный Markdown:**
```markdown
# Модуль DocumentSyncService
```

**Процесс в Word:**

1. **Создание параграфа:**
   ```
   Paragraph 1 создан
   Range: Start=0, End=0 (пустой)
   ```

2. **Вставка текста:**
   ```
   paragraph.Range.Text = "Модуль DocumentSyncService"
   Range: Start=0, End=25
   ```

3. **Применение форматирования:**
   ```
   textRange.Font.Size = 24
   textRange.Font.Bold = -1
   textRange.set_Style(wdStyleHeading1)
   ```

4. **Добавление переноса:**
   ```
   paragraph.Range.InsertParagraphAfter()
   Paragraph 2 создан (пустой)
   ```

**Результат в Word:**
```
[Paragraph 1] Модуль DocumentSyncService (24pt, жирный, стиль Heading 1)
[Paragraph 2] (пустой)
```

---

### Преобразование "## Описание"

**Исходный Markdown:**
```markdown
## Описание
```

**Процесс в Word:**

1. **Создание параграфа:**
   ```
   Paragraph 3 создан
   Range: Start=51, End=51 (пустой)
   ```

2. **Вставка текста:**
   ```
   paragraph.Range.Text = "Описание"
   Range: Start=51, End=59
   ```

3. **Применение форматирования:**
   ```
   textRange.Font.Size = 18
   textRange.Font.Bold = -1
   textRange.set_Style(wdStyleHeading2)
   ```

4. **Добавление переноса:**
   ```
   paragraph.Range.InsertParagraphAfter()
   Paragraph 4 создан (пустой)
   ```

**Результат в Word:**
```
[Paragraph 3] Описание (18pt, жирный, стиль Heading 2)
[Paragraph 4] (пустой)
```

---

## Ключевые концепции Word Interop

### 1. Range - это позиции, а не объекты

Важно понимать, что `Range` в Word - это не объект текста, а **указатель на позиции** в документе:

- `Range(0, 10)` - это не "первые 10 символов", а "диапазон от позиции 0 до позиции 10"
- При изменении документа позиции могут сдвигаться
- Range всегда актуален на момент его использования

### 2. Collapse - схлопывание Range

```csharp
range.Collapse(WdCollapseDirection.wdCollapseEnd);
```

- Схлопывает Range в точку (Start = End)
- Полезно для вставки в конкретную позицию

### 3. InsertAfter vs Text assignment

- `range.InsertAfter(text)` - вставляет текст после Range, Range расширяется
- `range.Text = text` - заменяет содержимое Range новым текстом

### 4. Стили Word

- Встроенные стили (`WdBuiltinStyle`) содержат предопределенное форматирование
- Применение стиля может переопределить некоторые свойства (размер, шрифт)
- Поэтому стиль применяется после установки размера шрифта (хотя стиль может его переопределить)

---

## Особенности реализации

### Обработка ошибок

Метод содержит обработку ошибок при применении стилей:

```csharp
try
{
    textRange.set_Style(headingStyle);
}
catch (Exception ex)
{
    textRange.set_Style(WdBuiltinStyle.wdStyleNormal);
}
```

Если стиль не может быть применен, используется обычный стиль Normal.

### Отладочные сообщения

Метод содержит множество `Debug.WriteLine()` для отслеживания процесса выполнения, что помогает при отладке.

### Поддержка форматированного текста

Метод поддерживает как простой текст, так и форматированный текст через `WordFormattedText`, что позволяет создавать заголовки с частичным форматированием (например, жирные и курсивные фрагменты).

---

## Заключение

Метод `ApplyToWord()` класса `WordTitle` выполняет следующие основные операции:

1. **Создает структуру** - новый параграф в документе
2. **Вставляет содержимое** - текст заголовка (простой или форматированный)
3. **Применяет форматирование** - размер шрифта, жирность, стиль заголовка
4. **Обеспечивает структуру** - добавляет перенос строки и очищает нумерацию

Все это происходит через работу с объектами `Range` и `Paragraph` в Word Interop API, которые представляют позиции и структуру документа соответственно.

