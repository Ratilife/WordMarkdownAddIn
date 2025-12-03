# Пошаговая инструкция для Visual Studio 2022
## Задача 1.1: Добавить переключение видимости Markdown/HTML
### Шаг 1: Добавить HTML-кнопки в BuildHtmlShell()
1. Откройте файл Controls/TaskPaneControl.cs
2. Найдите метод BuildHtmlShell() (около строки 383)
3. Внутри HTML-строки найдите <body> и элемент <div class="container">
4. Перед <div class="container"> добавьте панель кнопок:
```
<div class="view-controls">
    <button id="btn-split" class="view-btn active">Split</button>
    <button id="btn-markdown" class="view-btn">Markdown</button>
    <button id="btn-html" class="view-btn">HTML</button>
</div>

```
Важно: в C# строке с @"..." используйте двойные кавычки "" для экранирования.