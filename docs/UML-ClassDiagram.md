# UML Диаграмма Классов - WordMarkdownAddIn

## Описание

Этот документ содержит UML диаграмму классов проекта WordMarkdownAddIn в двух форматах:

1. **Mermaid** (в этом файле) - для отображения в GitHub, GitLab и других системах, поддерживающих Mermaid
2. **PlantUML** (файл `UML-ClassDiagram.puml`) - для использования в IDE (Visual Studio, IntelliJ IDEA) и онлайн-редакторах PlantUML

## Просмотр диаграмм

### Mermaid диаграмма
- Отображается автоматически в GitHub/GitLab
- Можно просмотреть в онлайн-редакторе: https://mermaid.live/
- Поддерживается многими Markdown просмотрщиками

### PlantUML диаграмма
- Используйте онлайн-редактор: http://www.plantuml.com/plantuml/uml/
- Или установите расширение для Visual Studio Code: "PlantUML"
- Или используйте плагин для Visual Studio

## Диаграмма классов проекта

```mermaid
classDiagram
    %% Основные классы приложения
    class ThisAddIn {
        +static CustomTaskPane MarkdownPane
        +static TaskPaneControl PaneControl
        +static MarkdownRibbon Ribbon
        +Dictionary~string,object~ Properties
        +static ThisAddIn Instance
        +Word.Application Application
        -CustomTaskPaneCollection CustomTaskPanes
        +TogglePane() void
        -ThisAddIn_Startup(sender, e) void
        -ThisAddIn_Shutdown(sender, e) void
        -Application_DocumentBeforeSave(Doc, SaveAsUI, Cancel) void
        -InternalStartup() void
    }

    class MarkdownRibbon {
        +RibbonTab tabMarkdown
        +RibbonGroup grpFile
        +RibbonButton btnSave
        +RibbonButton btnPanel
        +RibbonButton btnOpen
        +RibbonGroup grpFormat
        +RibbonButton bBold
        +RibbonButton bItalic
        +RibbonButton bStrike
        +RibbonButton bCode
        +RibbonGroup grpInsert
        +RibbonButton bH1
        +RibbonButton bH2
        +RibbonButton bH3
        +RibbonButton bList
        +RibbonButton bNumList
        +RibbonButton bCheckbox
        +RibbonButton bTable
        +RibbonButton bLink
        +RibbonButton bImage
        +RibbonButton bHR
        +RibbonButton bCodeBlock
        +RibbonButton bMermaid
        +RibbonButton bMath
        -MarkdownRibbon_Load(sender, e) void
        +btnSave_Click(sender, e) void
        +btnPanel_Click(sender, e) void
        +btnOpen_Click(sender, e) void
        +bBold_Click(sender, e) void
        +bItalic_Click(sender, e) void
        +bStrike_Click(sender, e) void
        +bH1_Click(sender, e) void
        +bH2_Click(sender, e) void
        +bList_Click(sender, e) void
        +bNumList_Click(sender, e) void
        +bCheckbox_Click(sender, e) void
        +bTable_Click(sender, e) void
        +bLink_Click(sender, e) void
        +bHR_Click(sender, e) void
        +bMermaid_Click(sender, e) void
        +bCodeBlock_Click(sender, e) void
        +bMath_Click(sender, e) void
    }

    class TaskPaneControl {
        -WebView2 _webView
        -MarkdownRenderService _renderer
        -string _latestMarkdown
        -bool _coreReady
        +TaskPaneControl()
        -OnLoadAsync(sender, e) void
        -CoreWebView2_WebMessageReceived(sender, e) void
        -PostRenderHtml(html) void
        +SetMarkdown(markdown) void
        +GetCachedMarkdown() string
        +GetMarkdownAsync() Task~string~
        -UnquoteJsonString(jsonQuoted) string
        +InsertInline(prefix, suffix) void
        +InsertSnippet(snippet) void
        +InsertHeading(level) void
        +InsertBulletList() void
        +InsertNumberedList() void
        +InsertCheckbox(isChecked) void
        +InsertTable(rows, cols) void
        +InsertLink(text, url) void
        +InsertImage(alt, path) void
        +InsertCodeBlock(language) void
        +InsertMermaidSample() void
        +InsertMermaid(mermaid_text) void
        +InsertMathSample() void
        +InsertMath(math_text) void
        +SaveMarkdownFile() void
        +OpenMarkdownFile() void
        -BuildHtmlShell() string
        -BuildHtmlShell_Old() string
    }

    class MarkdownRenderService {
        -MarkdownPipeline _pipeline
        -static Regex MermaidPreCodeRegex
        +MarkdownRenderService()
        +RenderoHtml(markdown) string
        -TransformMermaidBlocks(html) string
    }

    class DocumentSyncService {
        +const string NamespaceUri
        +LoadMarkdownFromActiveDocument(app) string
        +SaveMarkdownToActiveDocument(app, markdown) void
        -BuildXml(content) string
        -FindExistingPart(doc) CustomXMLPart
    }

    %% Внешние зависимости (Word/VSTO)
    class AddInBase {
        <<abstract>>
    }

    class RibbonBase {
        <<abstract>>
    }

    class UserControl {
        <<framework>>
    }

    class CustomTaskPane {
        <<external>>
        +bool Visible
        +int Width
        +MsoCTPDockPosition DockPosition
    }

    class WordApplication {
        <<external>>
        +Document ActiveDocument
        +DocumentBeforeSave event
    }

    class WebView2 {
        <<external>>
        +CoreWebView2 CoreWebView2
        +EnsureCoreWebView2Async() Task
        +ExecuteScriptAsync(script) Task
    }

    class MarkdownPipeline {
        <<external>>
    }

    class Globals {
        <<helper>>
        +static ThisAddIn ThisAddIn
        +static ApplicationFactory Factory
        +static ThisRibbonCollection Ribbons
    }

    %% Наследование
    ThisAddIn --|> AddInBase : extends
    MarkdownRibbon --|> RibbonBase : extends
    TaskPaneControl --|> UserControl : extends

    %% Композиция и агрегация
    ThisAddIn *-- CustomTaskPane : contains
    ThisAddIn *-- TaskPaneControl : creates
    ThisAddIn *-- MarkdownRibbon : creates
    ThisAddIn --> WordApplication : uses
    ThisAddIn --> DocumentSyncService : uses

    TaskPaneControl *-- WebView2 : contains
    TaskPaneControl *-- MarkdownRenderService : contains

    MarkdownRibbon --> ThisAddIn : calls static
    MarkdownRibbon --> TaskPaneControl : calls via ThisAddIn

    TaskPaneControl --> DocumentSyncService : uses
    TaskPaneControl --> MarkdownRenderService : uses

    MarkdownRenderService --> MarkdownPipeline : uses

    DocumentSyncService --> WordApplication : uses

    Globals --> ThisAddIn : references
```

## Описание классов

### ThisAddIn
Главный класс надстройки, наследуется от `AddInBase`. Управляет жизненным циклом приложения, создает и координирует работу всех компонентов.

**Ключевые связи:**
- Создает и управляет `CustomTaskPane` и `TaskPaneControl`
- Создает экземпляр `MarkdownRibbon`
- Использует `DocumentSyncService` для синхронизации с документами Word
- Связан с Word Application через событие `DocumentBeforeSave`

### MarkdownRibbon
Класс ленты интерфейса, наследуется от `RibbonBase`. Содержит все кнопки и обработчики событий для управления редактором Markdown.

**Ключевые связи:**
- Обращается к `TaskPaneControl` через статическое свойство `ThisAddIn.PaneControl`
- Вызывает методы `ThisAddIn.TogglePane()` для управления панелью

### TaskPaneControl
Пользовательский контрол, содержащий WebView2 редактор Markdown. Наследуется от `UserControl`. Обеспечивает все функции редактирования.

**Ключевые связи:**
- Содержит экземпляр `WebView2` для отображения HTML редактора
- Использует `MarkdownRenderService` для преобразования Markdown в HTML
- Использует `DocumentSyncService` для сохранения/загрузки Markdown

### MarkdownRenderService
Сервис для преобразования Markdown в HTML с использованием библиотеки Markdig.

**Ключевые связи:**
- Использует `MarkdownPipeline` из библиотеки Markdig
- Вызывается из `TaskPaneControl` для рендеринга предпросмотра

### DocumentSyncService
Статический сервис для синхронизации Markdown контента с Word документами через CustomXMLPart.

**Ключевые связи:**
- Использует Word Application для доступа к документам
- Работает с CustomXMLPart для хранения Markdown в документах

### Globals
Вспомогательный класс, автоматически сгенерированный VSTO. Предоставляет глобальный доступ к основным компонентам.

## Диаграмма зависимостей

```mermaid
graph TD
    A[Word Application] -->|события| B[ThisAddIn]
    B -->|создает| C[TaskPaneControl]
    B -->|создает| D[MarkdownRibbon]
    B -->|использует| E[DocumentSyncService]
    C -->|использует| F[MarkdownRenderService]
    C -->|использует| E
    C -->|содержит| G[WebView2]
    D -->|вызывает| C
    E -->|работает с| A
    F -->|использует| H[Markdig Pipeline]
```

## Основные паттерны

1. **Singleton**: `ThisAddIn.Instance` предоставляет единственный экземпляр надстройки
2. **Service Layer**: `DocumentSyncService` и `MarkdownRenderService` выделены в отдельные сервисы
3. **Observer**: Подписка на события Word (`DocumentBeforeSave`)
4. **Bridge**: Использование WebView2 для связи между C# и JavaScript
5. **Facade**: `ThisAddIn` предоставляет упрощенный интерфейс для доступа к компонентам

## Зависимости от внешних библиотек

- **Microsoft.Office.Tools**: Базовые классы VSTO (AddInBase, RibbonBase)
- **Microsoft.Web.WebView2**: Веб-браузер для редактора
- **Markdig**: Библиотека для преобразования Markdown в HTML
- **Microsoft.Office.Interop.Word**: COM-интерфейсы Word

