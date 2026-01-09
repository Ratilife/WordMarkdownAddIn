# Пошаговая инструкция: Реализация окна прогресса форматирования
## Visual Studio 2022

## Содержание
1. [Этап 1: Создание формы ProgressWindow](#этап-1-создание-формы-progresswindow)
2. [Этап 2: Создание класса FormattingProgressManager](#этап-2-создание-класса-formattingprogressmanager)
3. [Этап 3: Интеграция с WordMarkdownFormatter](#этап-3-интеграция-с-wordmarkdownformatter)
4. [Этап 4: Интеграция с MarkdownToWordFormatter](#этап-4-интеграция-с-markdowntowordformatter)
5. [Этап 5: Интеграция с WordToMarkdownService](#этап-5-интеграция-с-wordtomarkdownservice)
6. [Этап 6: Тестирование](#этап-6-тестирование)

---

## Этап 1: Создание формы ProgressWindow

### Шаг 1.1: Создание папки Forms

1. Откройте проект `WordMarkdownAddIn` в Visual Studio 2022
2. В **Solution Explorer** (Обозреватель решений) найдите корневую папку проекта
3. **Правой кнопкой мыши** на корневой папке проекта → **Add** → **New Folder**
4. Назовите папку `Forms`
5. Нажмите **Enter**

### Шаг 1.2: Добавление Windows Forms формы

1. **Правой кнопкой мыши** на папке `Forms` → **Add** → **Windows Form...**
2. В диалоговом окне **Add New Item**:
   - **Name:** `ProgressWindow.cs`
   - **Template:** Windows Form
3. Нажмите **Add**

### Шаг 1.3: Настройка свойств формы

1. В **Designer** (конструктор) выберите форму (кликните на пустое место формы)
2. В окне **Properties** (Свойства) установите следующие значения:

   | Свойство | Значение |
   |----------|----------|
   | `Name` | `ProgressWindow` |
   | `Text` | `Обработка...` |
   | `FormBorderStyle` | `FixedDialog` |
   | `MaximizeBox` | `False` |
   | `MinimizeBox` | `False` |
   | `ShowInTaskbar` | `False` |
   | `StartPosition` | `CenterScreen` |
   | `Size` | `400, 150` |
   | `TopMost` | `True` |

3. Нажмите **Enter** после каждого изменения

### Шаг 1.4: Добавление элементов управления

#### Шаг 1.4.1: Добавление Label для названия операции

1. В **Toolbox** (Панель элементов) найдите **Label**
2. Перетащите **Label** на форму
3. В **Properties** установите:
   - `Name`: `lblOperation`
   - `Text`: `Операция...`
   - `Font`: `Microsoft Sans Serif, 9pt, Style=Bold`
   - `Location`: `12, 12`
   - `Size`: `376, 20`
   - `AutoSize`: `False`

#### Шаг 1.4.2: Добавление ProgressBar

1. В **Toolbox** найдите **ProgressBar**
2. Перетащите **ProgressBar** на форму
3. В **Properties** установите:
   - `Name`: `progressBar`
   - `Location`: `12, 40`
   - `Size`: `376, 23`
   - `Style`: `Continuous`
   - `Minimum`: `0`
   - `Maximum`: `100`
   - `Value`: `0`

#### Шаг 1.4.3: Добавление Label для этапа

1. Перетащите еще один **Label** на форму
2. В **Properties** установите:
   - `Name`: `lblStage`
   - `Text`: `Инициализация...`
   - `Font`: `Microsoft Sans Serif, 8.25pt`
   - `Location`: `12, 75`
   - `Size`: `376, 20`
   - `AutoSize`: `False`

#### Шаг 1.4.4: Расположение элементов

Убедитесь, что элементы расположены следующим образом:

```
┌─────────────────────────────────────────┐
│  lblOperation (12, 12, 376x20)          │
│  progressBar (12, 40, 376x23)           │
│  lblStage (12, 75, 376x20)              │
└─────────────────────────────────────────┘
```

### Шаг 1.5: Редактирование кода формы

**Важно:** При создании Windows Form Visual Studio автоматически создает три файла:
- `ProgressWindow.cs` - основной код формы
- `ProgressWindow.Designer.cs` - код дизайнера (автоматически генерируется)
- `ProgressWindow.resx` - ресурсы формы (автоматически генерируется)

1. В **Solution Explorer** найдите `Forms/ProgressWindow.cs`
2. **Правой кнопкой мыши** → **View Code** (Просмотреть код)
3. Замените содержимое файла следующим кодом:

```csharp
using System;
using System.Windows.Forms;

namespace WordMarkdownAddIn.Forms
{
    /// <summary>
    /// Форма для отображения прогресса выполнения операции форматирования
    /// </summary>
    public partial class ProgressWindow : Form
    {
        /// <summary>
        /// Конструктор формы
        /// </summary>
        public ProgressWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Обновление прогресса и текста этапа
        /// </summary>
        /// <param name="value">Значение прогресса (0-100)</param>
        /// <param name="stage">Текст текущего этапа</param>
        public void UpdateProgress(int value, string stage)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int, string>(UpdateProgress), value, stage);
                return;
            }

            if (value < 0) value = 0;
            if (value > 100) value = 100;

            progressBar.Value = value;
            lblStage.Text = stage ?? "Обработка...";
            Application.DoEvents(); // Обновление UI
        }

        /// <summary>
        /// Установка названия операции
        /// </summary>
        /// <param name="operationName">Название операции</param>
        public void SetOperationName(string operationName)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(SetOperationName), operationName);
                return;
            }

            lblOperation.Text = operationName ?? "Обработка...";
            Text = operationName ?? "Обработка...";
        }

        /// <summary>
        /// Показ формы с обновлением UI
        /// </summary>
        public new void Show()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(Show));
                return;
            }

            base.Show();
            Application.DoEvents();
        }

        /// <summary>
        /// Скрытие формы
        /// </summary>
        public new void Hide()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(Hide));
                return;
            }

            base.Hide();
        }
    }
}
```

4. Сохраните файл (**Ctrl+S**)

### Шаг 1.6: Проверка компиляции

1. Нажмите **Build** → **Build Solution** (или **Ctrl+Shift+B**)
2. Убедитесь, что проект компилируется без ошибок
3. Если есть ошибки, проверьте:
   - Правильность имен элементов управления
   - Наличие всех using директив

---

## Этап 2: Создание класса FormattingProgressManager

### Шаг 2.1: Создание файла FormattingProgressManager.cs

1. В **Solution Explorer** найдите папку `Services`
2. **Правой кнопкой мыши** на папке `Services` → **Add** → **Class...**
3. В диалоговом окне:
   - **Name:** `FormattingProgressManager.cs`
4. Нажмите **Add**

### Шаг 2.2: Реализация класса FormattingProgressManager

1. Откройте файл `Services/FormattingProgressManager.cs`
2. Замените содержимое следующим кодом:

```csharp
using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using WordMarkdownAddIn.Forms;

namespace WordMarkdownAddIn.Services
{
    /// <summary>
    /// Менеджер для управления отображением окна прогресса при форматировании
    /// </summary>
    public class FormattingProgressManager : IDisposable
    {
        private ProgressWindow _progressWindow;
        private Stopwatch _stopwatch;
        private readonly TimeSpan _threshold;
        private bool _isProgressVisible;
        private Thread _uiThread;
        private readonly object _lockObject = new object();

        /// <summary>
        /// Порог времени для показа окна прогресса (по умолчанию 7 секунд)
        /// </summary>
        public TimeSpan Threshold => _threshold;

        /// <summary>
        /// Флаг, указывающий, видно ли окно прогресса
        /// </summary>
        public bool IsProgressVisible
        {
            get
            {
                lock (_lockObject)
                {
                    return _isProgressVisible;
                }
            }
        }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="thresholdSeconds">Порог времени в секундах (по умолчанию 7)</param>
        public FormattingProgressManager(int thresholdSeconds = 7)
        {
            _threshold = TimeSpan.FromSeconds(thresholdSeconds);
            _stopwatch = new Stopwatch();
        }

        /// <summary>
        /// Начало операции
        /// </summary>
        /// <param name="operationName">Название операции</param>
        public void StartOperation(string operationName)
        {
            lock (_lockObject)
            {
                _isProgressVisible = false;
                _stopwatch.Restart();

                // Запускаем таймер для проверки порога
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = (int)_threshold.TotalMilliseconds;
                timer.Tick += (sender, e) =>
                {
                    timer.Stop();
                    timer.Dispose();

                    if (_stopwatch.IsRunning && !_isProgressVisible)
                    {
                        ShowProgressWindow(operationName);
                    }
                };
                timer.Start();
            }
        }

        /// <summary>
        /// Показ окна прогресса
        /// </summary>
        private void ShowProgressWindow(string operationName)
        {
            lock (_lockObject)
            {
                if (_isProgressVisible)
                    return;

                try
                {
                    // Создаем окно в UI потоке
                    if (Application.MessageLoop)
                    {
                        // Мы в UI потоке
                        _progressWindow = new ProgressWindow();
                        _progressWindow.SetOperationName(operationName);
                        _progressWindow.Show();
                        _isProgressVisible = true;
                    }
                    else
                    {
                        // Запускаем в отдельном потоке (для безопасности)
                        _uiThread = new Thread(() =>
                        {
                            _progressWindow = new ProgressWindow();
                            _progressWindow.SetOperationName(operationName);
                            Application.Run(_progressWindow);
                        });
                        _uiThread.SetApartmentState(ApartmentState.STA);
                        _uiThread.Start();
                        _isProgressVisible = true;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[FormattingProgressManager] Ошибка при показе окна: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Обновление прогресса
        /// </summary>
        /// <param name="value">Значение прогресса (0-100)</param>
        /// <param name="stage">Текст текущего этапа</param>
        public void UpdateProgress(int value, string stage)
        {
            lock (_lockObject)
            {
                // Если прошло больше порога, показываем окно
                if (!_isProgressVisible && _stopwatch.Elapsed >= _threshold)
                {
                    // Окно должно было показаться по таймеру, но на всякий случай показываем здесь
                    if (_progressWindow == null)
                    {
                        ShowProgressWindow("Обработка...");
                    }
                }

                if (_isProgressVisible && _progressWindow != null)
                {
                    try
                    {
                        _progressWindow.UpdateProgress(value, stage);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[FormattingProgressManager] Ошибка при обновлении прогресса: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Завершение операции
        /// </summary>
        public void CompleteOperation()
        {
            lock (_lockObject)
            {
                _stopwatch.Stop();

                if (_isProgressVisible && _progressWindow != null)
                {
                    try
                    {
                        if (_progressWindow.InvokeRequired)
                        {
                            _progressWindow.Invoke(new Action(() =>
                            {
                                _progressWindow.Hide();
                                _progressWindow.Close();
                                _progressWindow.Dispose();
                            }));
                        }
                        else
                        {
                            _progressWindow.Hide();
                            _progressWindow.Close();
                            _progressWindow.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[FormattingProgressManager] Ошибка при закрытии окна: {ex.Message}");
                    }
                    finally
                    {
                        _progressWindow = null;
                        _isProgressVisible = false;
                    }
                }
            }
        }

        /// <summary>
        /// Освобождение ресурсов
        /// </summary>
        public void Dispose()
        {
            CompleteOperation();

            if (_stopwatch != null)
            {
                _stopwatch.Stop();
                _stopwatch = null;
            }
        }
    }
}
```

3. Сохраните файл (**Ctrl+S**)

### Шаг 2.3: Проверка компиляции

1. Нажмите **Build** → **Build Solution** (или **Ctrl+Shift+B**)
2. Убедитесь, что проект компилируется без ошибок
3. Если есть ошибки, проверьте:
   - Правильность using директив
   - Наличие ссылки на Forms.ProgressWindow

---

## Этап 3: Интеграция с WordMarkdownFormatter

### Шаг 3.1: Добавление поля FormattingProgressManager

1. Откройте файл `Services/WordMarkdownFormatter.cs`
2. Найдите класс `WordMarkdownFormatter` (около строки 1194)
3. Найдите приватные поля класса (после строки 1206)
4. Добавьте новое поле:

```csharp
private FormattingProgressManager _progressManager;
```

### Шаг 3.2: Инициализация FormattingProgressManager в конструкторе

1. Найдите конструктор `WordMarkdownFormatter()` (около строки 1211)
2. В конце конструктора (перед закрывающей скобкой) добавьте:

```csharp
_progressManager = new FormattingProgressManager(7); // 7 секунд порог
```

### Шаг 3.3: Модификация метода FormatMarkdownInWord

1. Найдите метод `FormatMarkdownInWord(Range targetRange = null)` (около строки 1246)
2. В начале метода, после проверки `_isProcessing`, добавьте:

```csharp
_progressManager?.StartOperation("Форматирование Markdown в Word");
```

3. После валидации документа (после строки 1262) добавьте:

```csharp
_progressManager?.UpdateProgress(10, "Проверка документа...");
```

4. После определения диапазона обработки (после строки 1271) добавьте:

```csharp
_progressManager?.UpdateProgress(15, "Подготовка к обработке...");
```

5. После поиска элементов Markdown (после строки 1282) добавьте:

```csharp
if (elements != null && elements.Count > 0)
{
    _progressManager?.UpdateProgress(30, $"Найдено {elements.Count} элементов Markdown...");
}
```

6. После сортировки элементов (после строки 1295) добавьте:

```csharp
_progressManager?.UpdateProgress(40, "Сортировка элементов...");
```

7. В цикле обработки элементов (около строки 1300) замените цикл:

**Было:**
```csharp
foreach (var element in sortedElements.OrderByDescending(e => e.StartPosition))
{
    ApplyFormattingToElement(element, rangeToProcess);
}
```

**Стало:**
```csharp
int processed = 0;
int total = sortedElements.Count;
foreach (var element in sortedElements.OrderByDescending(e => e.StartPosition))
{
    ApplyFormattingToElement(element, rangeToProcess);
    processed++;
    
    // Обновляем прогресс: 40-90% для обработки элементов
    int progress = 40 + (int)((double)processed / total * 50);
    _progressManager?.UpdateProgress(progress, $"Обработано {processed} из {total} элементов...");
}
```

8. Перед строкой `Debug.WriteLine($"[FormatMarkdownInWord] Обработано...")` добавьте:

```csharp
_progressManager?.UpdateProgress(100, "Завершено");
```

9. В блоке `finally` (около строки 1311) перед `ResetProcessingState()` добавьте:

```csharp
_progressManager?.CompleteOperation();
```

### Шаг 3.4: Добавление using директивы

1. В начале файла `WordMarkdownFormatter.cs` найдите секцию using
2. Убедитесь, что есть:
```csharp
using WordMarkdownAddIn.Services;
```

Если нет, добавьте.

### Шаг 3.5: Проверка компиляции

1. Нажмите **Build** → **Build Solution** (или **Ctrl+Shift+B**)
2. Убедитесь, что проект компилируется без ошибок

---

## Этап 4: Интеграция с MarkdownToWordFormatter

### Шаг 4.1: Добавление поля FormattingProgressManager

1. Откройте файл `Services/MarkdownToWordFormatter.cs`
2. Найдите класс `MarkdownToWordFormatter` (около строки 18)
3. Найдите приватные поля класса (после строки 22)
4. Добавьте новое поле:

```csharp
private FormattingProgressManager _progressManager;
```

### Шаг 4.2: Инициализация в конструкторе

1. Найдите конструктор `MarkdownToWordFormatter()` (около строки 24)
2. В конце конструктора добавьте:

```csharp
_progressManager = new FormattingProgressManager(7);
```

### Шаг 4.3: Модификация метода ApplyMarkdownToWord

1. Найдите метод `ApplyMarkdownToWord(string markdown)` в файле
2. В начале метода добавьте:

```csharp
_progressManager?.StartOperation("Применение Markdown к Word");
```

3. Найдите ключевые этапы обработки в методе и добавьте вызовы `UpdateProgress`:

```csharp
_progressManager?.UpdateProgress(10, "Парсинг Markdown...");
// ... после парсинга

_progressManager?.UpdateProgress(30, "Преобразование элементов...");
// ... после преобразования

_progressManager?.UpdateProgress(60, "Применение форматирования...");
// ... во время применения

_progressManager?.UpdateProgress(100, "Завершено");
// ... перед завершением
```

4. В конце метода (в блоке `finally` или перед `return`) добавьте:

```csharp
_progressManager?.CompleteOperation();
```

**Примечание:** Если метод `ApplyMarkdownToWord` большой, добавьте вызовы `UpdateProgress` на основных этапах обработки (парсинг, преобразование, применение).

### Шаг 4.4: Проверка компиляции

1. Нажмите **Build** → **Build Solution**
2. Убедитесь, что проект компилируется без ошибок

---

## Этап 5: Интеграция с WordToMarkdownService

### Шаг 5.1: Добавление поля FormattingProgressManager

1. Откройте файл `Services/WordToMarkdownService.cs`
2. Найдите класс `WordToMarkdownService` (около строки 15)
3. Найдите приватные поля класса (после строки 30)
4. Добавьте новое поле:

```csharp
private FormattingProgressManager _progressManager;
```

### Шаг 5.2: Инициализация в конструкторе

1. Найдите конструктор `WordToMarkdownService()` (около строки 31)
2. В конце конструктора добавьте:

```csharp
_progressManager = new FormattingProgressManager(7);
```

### Шаг 5.3: Модификация метода ConvertToMarkdown

1. Найдите метод `ConvertToMarkdown()` (около строки 83)
2. В начале метода (после `try {`) добавьте:

```csharp
_progressManager?.StartOperation("Преобразование Word в Markdown");
```

3. После извлечения структуры документа (после строки 88) добавьте:

```csharp
_progressManager?.UpdateProgress(20, $"Извлечено {elements.Count} элементов...");
```

4. В цикле преобразования элементов (около строки 96) добавьте обновление прогресса:

**Найдите цикл:**
```csharp
for (int i = 0; i < elements.Count; i++)
{
    var element = elements[i];
    // ... обработка
}
```

**Добавьте обновление прогресса:**
```csharp
for (int i = 0; i < elements.Count; i++)
{
    var element = elements[i];
    // ... существующая обработка ...
    
    // Обновление прогресса: 20-90% для преобразования
    int progress = 20 + (int)((double)(i + 1) / elements.Count * 70);
    _progressManager?.UpdateProgress(progress, $"Преобразовано {i + 1} из {elements.Count} элементов...");
}
```

5. Перед `return sb.ToString();` добавьте:

```csharp
_progressManager?.UpdateProgress(100, "Завершено");
```

6. В блоке `catch` или в конце метода добавьте:

```csharp
_progressManager?.CompleteOperation();
```

### Шаг 5.4: Проверка компиляции

1. Нажмите **Build** → **Build Solution**
2. Убедитесь, что проект компилируется без ошибок

---

## Этап 6: Тестирование

### Шаг 6.1: Подготовка тестового документа

1. Откройте Microsoft Word
2. Создайте новый документ
3. Вставьте большой объем текста с Markdown-синтаксисом (например, скопируйте текст несколько раз):

```
# Заголовок 1

## Заголовок 2

Это **жирный текст** и это *курсивный текст*.

- Элемент списка 1
- Элемент списка 2
- Элемент списка 3

[Ссылка](https://example.com)

`инлайн-код`

> Цитата

---

```python
def hello():
    print("Hello, World!")
```
```

4. Повторите вставку несколько раз, чтобы создать большой документ (более 1000 элементов)

### Шаг 6.2: Тестирование быстрой операции

1. Создайте небольшой документ с несколькими элементами Markdown
2. Запустите проект (**F5**)
3. В Word нажмите кнопку **"Форматировать Markdown"** в ленте
4. **Ожидаемый результат:** Операция завершается быстро, окно прогресса не появляется

### Шаг 6.3: Тестирование долгой операции

1. Используйте большой тестовый документ из шага 6.1
2. Запустите проект (**F5**)
3. В Word нажмите кнопку **"Форматировать Markdown"** в ленте
4. Подождите 7 секунд
5. **Ожидаемый результат:** 
   - Через 7 секунд появляется окно прогресса
   - Индикатор прогресса обновляется
   - Текст этапа меняется
   - Окно закрывается после завершения

### Шаг 6.4: Тестирование других операций

1. **Тест "Применить к Word":**
   - Откройте панель Markdown
   - Вставьте большой Markdown-текст
   - Нажмите кнопку **"Применить к Word"**
   - Проверьте появление окна прогресса

2. **Тест "Преобразовать в Markdown":**
   - Откройте большой документ Word
   - Нажмите кнопку **"Преобразовать в Markdown"**
   - Проверьте появление окна прогресса

### Шаг 6.5: Проверка обработки ошибок

1. Попробуйте выполнить операцию на защищенном документе
2. **Ожидаемый результат:** Окно прогресса корректно закрывается даже при ошибке

### Шаг 6.6: Проверка производительности

1. Запустите операцию форматирования
2. Убедитесь, что производительность не ухудшилась значительно
3. Проверьте, что нет утечек памяти (запустите несколько операций подряд)

---

## Дополнительные замечания

### Отладка

Если окно прогресса не появляется:

1. Проверьте, что `_progressManager` не равен `null`
2. Проверьте логи в **Output** окне Visual Studio (View → Output)
3. Убедитесь, что операция действительно длится более 7 секунд

### Улучшения (опционально)

В будущем можно добавить:

1. Кнопку "Отмена" в окне прогресса
2. Настройку порога времени в настройках приложения
3. Более детальную информацию о прогрессе (скорость, оставшееся время)

---

## Заключение

После выполнения всех этапов окно прогресса будет автоматически появляться при операциях форматирования, которые длятся более 7 секунд. Это улучшит пользовательский опыт, предоставляя обратную связь о ходе выполнения длительных операций.

**Дата создания:** 2024  
**Версия:** 1.0

