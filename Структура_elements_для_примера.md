# Структура переменной `elements` для примера Markdown

## Исходный Markdown текст

```markdown
# Модуль DocumentSyncService

## Описание
Сервис для синхронизации Markdown контента с Word документом через CustomXMLPart. Обеспечивает сохранение и загрузку Markdown данных в/из Word документов, что позволяет сохранять исходный Markdown вместе с документом.

## Что реализовано

### ✅ Сохранение Markdown в документ
- Метод `SaveMarkdownToActiveDocument()` - сохраняет Markdown в активный документ
- Использует CustomXMLPart для хранения данных
- Автоматически удаляет старую версию перед сохранением новой
- Оборачивает содержимое в CDATA для защиты специальных XML символов
```

## Структура переменной `elements`

После парсинга Markdown через `Markdown.Parse()` и обработки блоков методом `ProcessBlock()`, переменная `elements` будет содержать **9 элементов** типа `List<IWordElement>`:

---

### Элемент #1: `WordTitle` (уровень 1)

```csharp
new WordTitle(
    Text: "",
    Content: new WordFormattedText {
        Runs = new List<FormattedRun> {
            new FormattedRun {
                Text: "Модуль DocumentSyncService",
                IsBold: false,
                IsItalic: false,
                // ... остальные свойства форматирования = false
            }
        }
    },
    Level: 1
)
```

**Свойства:**
- `ElementType`: `"Title"`
- `Level`: `1`
- `Content.Runs[0].Text`: `"Модуль DocumentSyncService"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом "Модуль DocumentSyncService"
- Применяется стиль `WdBuiltinStyle.wdStyleHeading1`
- Устанавливается размер шрифта: `24pt`
- Текст делается жирным

---

### Элемент #2: `WordTitle` (уровень 2)

```csharp
new WordTitle(
    Text: "",
    Content: new WordFormattedText {
        Runs = new List<FormattedRun> {
            new FormattedRun {
                Text: "Описание",
                IsBold: false,
                IsItalic: false,
                // ... остальные свойства форматирования = false
            }
        }
    },
    Level: 2
)
```

**Свойства:**
- `ElementType`: `"Title"`
- `Level`: `2`
- `Content.Runs[0].Text`: `"Описание"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом "Описание"
- Применяется стиль `WdBuiltinStyle.wdStyleHeading2`
- Устанавливается размер шрифта: `18pt`
- Текст делается жирным

---

### Элемент #3: `WordParagraph`

```csharp
new WordParagraph(
    StyleName: "Normal", // или текущий стиль пользователя
    Content: new WordFormattedText {
        Runs = new List<FormattedRun> {
            new FormattedRun {
                Text: "Сервис для синхронизации Markdown контента с Word документом через CustomXMLPart. Обеспечивает сохранение и загрузку Markdown данных в/из Word документов, что позволяет сохранять исходный Markdown вместе с документом.",
                IsBold: false,
                IsItalic: false,
                // ... остальные свойства форматирования = false
            }
        }
    }
)
```

**Свойства:**
- `ElementType`: `"Paragraph"`
- `StyleName`: `"Normal"` (или значение из `GetCurrentParagraphStyle()`)
- `Content.Runs[0].Text`: Полный текст параграфа

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с обычным текстом
- Применяется стиль `WdBuiltinStyle.wdStyleNormal`
- Текст вставляется без дополнительного форматирования

---

### Элемент #4: `WordTitle` (уровень 2)

```csharp
new WordTitle(
    Text: "",
    Content: new WordFormattedText {
        Runs = new List<FormattedRun> {
            new FormattedRun {
                Text: "Что реализовано",
                IsBold: false,
                IsItalic: false,
                // ... остальные свойства форматирования = false
            }
        }
    },
    Level: 2
)
```

**Свойства:**
- `ElementType`: `"Title"`
- `Level`: `2`
- `Content.Runs[0].Text`: `"Что реализовано"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом "Что реализовано"
- Применяется стиль `WdBuiltinStyle.wdStyleHeading2`
- Устанавливается размер шрифта: `18pt`
- Текст делается жирным

---

### Элемент #5: `WordTitle` (уровень 3)

```csharp
new WordTitle(
    Text: "",
    Content: new WordFormattedText {
        Runs = new List<FormattedRun> {
            new FormattedRun {
                Text: "✅ Сохранение Markdown в документ",
                IsBold: false,
                IsItalic: false,
                // ... остальные свойства форматирования = false
            }
        }
    },
    Level: 3
)
```

**Свойства:**
- `ElementType`: `"Title"`
- `Level`: `3`
- `Content.Runs[0].Text`: `"✅ Сохранение Markdown в документ"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом "✅ Сохранение Markdown в документ"
- Применяется стиль `WdBuiltinStyle.wdStyleHeading3`
- Устанавливается размер шрифта: `14pt`
- Текст делается жирным

---

### Элемент #6: `WordListItem` (первый элемент списка)

```csharp
new WordListItem(
    Contents: new List<WordFormattedText> {
        new WordFormattedText {
            Runs = new List<FormattedRun> {
                // Run #1: "Метод "
                new FormattedRun {
                    Text: "Метод ",
                    IsBold: false,
                    IsItalic: false,
                    // ...
                },
                // Run #2: "SaveMarkdownToActiveDocument()" (код, моноширинный)
                new FormattedRun {
                    Text: "SaveMarkdownToActiveDocument()",
                    IsBold: false,
                    IsItalic: false,
                    // Возможно, имеет моноширинный шрифт (если обрабатывается как код)
                    // ...
                },
                // Run #3: " - сохраняет Markdown в активный документ"
                new FormattedRun {
                    Text: " - сохраняет Markdown в активный документ",
                    IsBold: false,
                    IsItalic: false,
                    // ...
                }
            }
        }
    },
    IsOrdered: false  // маркированный список
)
```

**Свойства:**
- `ElementType`: `"ListItem"`
- `IsOrdered`: `false` (маркированный список)
- `Contents[0].Runs.Count`: `3` (согласно логам: "Завершено, создано Runs: 3")
- `Contents[0].Runs[0].Text`: `"Метод "`
- `Contents[0].Runs[1].Text`: `"SaveMarkdownToActiveDocument()"` (внутри обратных кавычек)
- `Contents[0].Runs[2].Text`: `" - сохраняет Markdown в активный документ"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф
- Вставляется форматированный текст (с обработкой кода в обратных кавычках)
- Применяется форматирование маркированного списка: `listParagraph.Range.ListFormat.ApplyBulletDefault()`

---

### Элемент #7: `WordListItem` (второй элемент списка)

```csharp
new WordListItem(
    Contents: new List<WordFormattedText> {
        new WordFormattedText {
            Runs = new List<FormattedRun> {
                new FormattedRun {
                    Text: "Использует CustomXMLPart для хранения данных",
                    IsBold: false,
                    IsItalic: false,
                    // ...
                }
            }
        }
    },
    IsOrdered: false
)
```

**Свойства:**
- `ElementType`: `"ListItem"`
- `IsOrdered`: `false`
- `Contents[0].Runs[0].Text`: `"Использует CustomXMLPart для хранения данных"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом
- Применяется форматирование маркированного списка

---

### Элемент #8: `WordListItem` (третий элемент списка)

```csharp
new WordListItem(
    Contents: new List<WordFormattedText> {
        new WordFormattedText {
            Runs = new List<FormattedRun> {
                new FormattedRun {
                    Text: "Автоматически удаляет старую версию перед сохранением новой",
                    IsBold: false,
                    IsItalic: false,
                    // ...
                }
            }
        }
    },
    IsOrdered: false
)
```

**Свойства:**
- `ElementType`: `"ListItem"`
- `IsOrdered`: `false`
- `Contents[0].Runs[0].Text`: `"Автоматически удаляет старую версию перед сохранением новой"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом
- Применяется форматирование маркированного списка

---

### Элемент #9: `WordListItem` (четвертый элемент списка)

```csharp
new WordListItem(
    Contents: new List<WordFormattedText> {
        new WordFormattedText {
            Runs = new List<FormattedRun> {
                new FormattedRun {
                    Text: "Оборачивает содержимое в CDATA для защиты специальных XML символов",
                    IsBold: false,
                    IsItalic: false,
                    // ...
                }
            }
        }
    },
    IsOrdered: false
)
```

**Свойства:**
- `ElementType`: `"ListItem"`
- `IsOrdered`: `false`
- `Contents[0].Runs[0].Text`: `"Оборачивает содержимое в CDATA для защиты специальных XML символов"`

**При вызове `element.ApplyToWord(_activeDoc)`:**
- Создается параграф с текстом
- Применяется форматирование маркированного списка

---

## Процесс применения элементов к Word

В методе `ApplyMarkdownToWord()` на строке **680** происходит вызов:

```639:680:Services/MarkdownToWordFormatter.cs
// 2. Преобразуем в коллекцию IWordElement
var elements = new List<IWordElement>();
int blockIndex = 0;
foreach(var block in document)
{
    blockIndex++;
    string blockType = block?.GetType().Name ?? "null";
    System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Блок #{blockIndex}: {blockType}");

    // Списки обрабатываем отдельно, так как они могут содержать несколько элементов
    if (block is ListBlock listBlock)
    {
        var listItems = ProcessListItems(listBlock);
        System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Добавлено элементов списка: {listItems.Count}");
        elements.AddRange(listItems);
    }
    else
    {
        IWordElement element = ProcessBlock(block);
        if(element != null)
        {
            System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Элемент добавлен: {element.GetType().Name}, ElementType: {element.ElementType}");
            elements.Add(element);
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] ВНИМАНИЕ: ProcessBlock вернул null для блока типа {blockType}");
        }
    }
}

System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Всего элементов создано: {elements.Count}");

// 3. Применяем все элементы к Word
int elementIndex = 0;
foreach(var element in elements)
{
    elementIndex++;
    System.Diagnostics.Debug.WriteLine($"[ApplyMarkdownToWord] Применение элемента #{elementIndex}: {element.GetType().Name}, ElementType: {element.ElementType}");
    try
    {
        element.ApplyToWord(_activeDoc);
```

### Последовательность вызовов:

1. **Элемент #1** (`WordTitle`, Level=1) → `wordTitle.ApplyToWord(_activeDoc)`
   - Вызывается метод `WordTitle.ApplyToWord()` (строка 836 в `IWordElement.cs`)
   - Создается параграф, вставляется текст, применяется стиль Heading 1

2. **Элемент #2** (`WordTitle`, Level=2) → `wordTitle.ApplyToWord(_activeDoc)`
   - Вызывается метод `WordTitle.ApplyToWord()`
   - Создается параграф, вставляется текст, применяется стиль Heading 2

3. **Элемент #3** (`WordParagraph`) → `wordParagraph.ApplyToWord(_activeDoc)`
   - Вызывается метод `WordParagraph.ApplyToWord()` (строка 583 в `IWordElement.cs`)
   - Создается параграф, вставляется текст, применяется стиль Normal

4. **Элемент #4** (`WordTitle`, Level=2) → `wordTitle.ApplyToWord(_activeDoc)`
   - Вызывается метод `WordTitle.ApplyToWord()`
   - Создается параграф, вставляется текст, применяется стиль Heading 2

5. **Элемент #5** (`WordTitle`, Level=3) → `wordTitle.ApplyToWord(_activeDoc)`
   - Вызывается метод `WordTitle.ApplyToWord()`
   - Создается параграф, вставляется текст, применяется стиль Heading 3

6. **Элемент #6** (`WordListItem`) → `wordListItem.ApplyToWord(_activeDoc)`
   - Вызывается метод `WordListItem.ApplyToWord()` (строка 461 в `IWordElement.cs`)
   - Создается параграф, вставляется форматированный текст, применяется маркированный список

7. **Элемент #7** (`WordListItem`) → `wordListItem.ApplyToWord(_activeDoc)`
   - Аналогично элементу #6

8. **Элемент #8** (`WordListItem`) → `wordListItem.ApplyToWord(_activeDoc)`
   - Аналогично элементу #6

9. **Элемент #9** (`WordListItem`) → `wordListItem.ApplyToWord(_activeDoc)`
   - Аналогично элементу #6

## Важные детали из логов отладки

### Обработка первого элемента списка:

Из логов видно, что первый элемент списка содержит **3 Runs**:
```
[ConvertInlineToWordFormattedText] Завершено, создано Runs: 3
[ConvertInlineToWordFormattedText] Общий текст: 'Метод SaveMarkdownToActiveDocument() - сохраняет Markdown в активный документ'
```

Это происходит потому, что текст `Метод `SaveMarkdownToActiveDocument()` - сохраняет...` содержит:
1. Обычный текст: `"Метод "`
2. Код в обратных кавычках: `"SaveMarkdownToActiveDocument()"` (может иметь моноширинный шрифт)
3. Обычный текст: `" - сохраняет Markdown в активный документ"`

### Игнорирование LinkReferenceDefinitionGroup:

Из логов видно, что блок #7 (`LinkReferenceDefinitionGroup`) был проигнорирован:
```
[ApplyMarkdownToWord] Блок #7: LinkReferenceDefinitionGroup
[ProcessBlock] Неизвестный тип блока: LinkReferenceDefinitionGroup
[ApplyMarkdownToWord] ВНИМАНИЕ: ProcessBlock вернул null для блока типа LinkReferenceDefinitionGroup
```

Это нормально - Markdig может создавать такие блоки для внутренних ссылок, но они не влияют на визуальное отображение.

## Итоговая структура `elements`

```csharp
List<IWordElement> elements = new List<IWordElement>
{
    // #1
    new WordTitle("", formattedText1, 1),
    
    // #2
    new WordTitle("", formattedText2, 2),
    
    // #3
    new WordParagraph("Normal", formattedText3),
    
    // #4
    new WordTitle("", formattedText4, 2),
    
    // #5
    new WordTitle("", formattedText5, 3),
    
    // #6
    new WordListItem(listItemContents1, false),
    
    // #7
    new WordListItem(listItemContents2, false),
    
    // #8
    new WordListItem(listItemContents3, false),
    
    // #9
    new WordListItem(listItemContents4, false)
};
```

Все эти элементы последовательно применяются к документу Word через вызов `element.ApplyToWord(_activeDoc)` в цикле на строке 680.

