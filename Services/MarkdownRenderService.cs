using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Markdig;

namespace WordMarkdownAddIn.Services
{
    public class MarkdownRenderService
    {
        private readonly MarkdownPipeline _pipeline;                            // Содержит все настройки преобразования Markdown → HTML
                                                                                // Переменная может быть присвоена только в конструкторе
                                                                                // Не может быть изменена после инициализации
                                                                                // Гарантирует безопасность потоков (thread-safe) для неизменяемых данных
        
        private static readonly Regex MermaidPreCodeRegex = new Regex(          //MermaidPreCodeRegex -это предкомпилированное регулярное выражение для поиска и извлечения Mermaid-блоков из HTML.
            "<pre><code class=\"language-mermaid\">([\\s\\S]*?)</code></pre>",  //Ищет блоки кода с классом language-mermaid
            RegexOptions.Compiled | RegexOptions.IgnoreCase);                   //([\s\S]*?) - захватывает любое содержимое (включая переносы строк)
                                                                                //*? - нежадный квантификатор (захватывает минимальное совпадение)
                                                                                //RegexOptions.Compiled
                                                                                //Компилирует регулярное выражение в IL-код
                                                                                //Увеличивает скорость выполнения в 10-100 раз
                                                                                //Увеличивает потребление памяти (однократно при загрузке класса)
                                                                                //RegexOptions.IgnoreCase
                                                                                //Делает поиск нечувствительным к регистру
                                                                                //Почему static readonly:
                                                                                //static - Регулярное выражение не зависит от экземпляра класса; дин экземпляр на весь класс (не создается для каждого объекта)
                                                                                //readonly - Не может быть изменено после инициализации;Гарантирует безопасность потоков
                                                                                
        public MarkdownRenderService() 
        {
            _pipeline = new MarkdownPipelineBuilder()                           // 1. Создание строителя конвейера Markdig
                .UseAdvancedExtensions()                                        // 2. Включение расширенных расширений Markdig
                .UsePipeTables()                                                // 3. Включение поддержки таблиц с использованием символа pipe (|)
                .UseTaskLists()                                                 // 4. Включение поддержки списков задач (чекбоксов)
                .UseMathematics()                                               // 5. Включение поддержки математических формул (LaTeX)
                .Build();                                                       // 6. Финальная сборка конвейера
        }
        
        public string RenderoHtml(string markdown) 
        {
            markdown = markdown ?? string.Empty;                                // 1. Проверка на null и установка значения по умолчанию
            var html = Markdown.ToHtml(markdown, _pipeline);                    // 2. Преобразование Markdown в HTML с использованием настроенного конвейера
            html = TransformMermaidBlocks(html);                                // 3. Специальная обработка Mermaid-диаграмм
            return html;                                                        // 4. Возврат конечного HTML-результата
        }

        private static string TransformMermaidBlocks(string html)
        {
            return MermaidPreCodeRegex.Replace(html, m =>                       // 1. Поиск и замена Mermaid-блоков с помощью регулярного выражения
            {
                var inner = m.Groups[1].Value;                                  // 2. Извлечение содержимого из первой группы захвата
                var decoded = WebUtility.HtmlDecode(inner);                     // 3. Декодирование HTML-сущностей
                return "<div class=\"mermaid\">" + decoded + "</div>";          // 4. Формирование нового HTML-блока для Mermaid
            });
        }
    }
}
