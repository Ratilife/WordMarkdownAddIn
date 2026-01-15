using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordMarkdownAddIn.Services
{
    public class MermaidExportService
    {
        private static readonly Regex MermaidBlockRegex = new Regex(
            @"```mermaid\s*\n([\s\S]*?)\n```",
            RegexOptions.Compiled | RegexOptions.Multiline
        );

        public List<string> ExtractMermaidDiagrams(string markdown)
        {
            var diagrams = new List<string>();
            
            if (string.IsNullOrEmpty(markdown))
            {
                return diagrams;
            }

            var matches = MermaidBlockRegex.Matches(markdown);
            
            foreach (Match match in matches)
            {
                if (match.Groups.Count > 1)
                {
                    string diagramCode = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(diagramCode))
                    {
                        diagrams.Add(diagramCode);
                    }
                }
            }
            
            return diagrams;
        }

        private async Task<T> InvokeOnUIThreadAsync<T>(WebView2 webView, Func<CoreWebView2, Task<T>> action)
        {
            if (webView == null)
                throw new Exception("WebView2 не инициализирован");

            if (webView.InvokeRequired)
            {
                var tcs = new TaskCompletionSource<T>();
                webView.BeginInvoke(new Action(async () =>
                {
                    try
                    {
                        if (webView.CoreWebView2 == null)
                            throw new Exception("CoreWebView2 не инициализирован");
                        var result = await action(webView.CoreWebView2);
                        tcs.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }));
                return await tcs.Task;
            }
            else
            {
                if (webView.CoreWebView2 == null)
                    throw new Exception("CoreWebView2 не инициализирован");
                return await action(webView.CoreWebView2);
            }
        }

        private async Task InvokeOnUIThreadAsync(WebView2 webView, Func<CoreWebView2, Task> action)
        {
            if (webView == null)
                throw new Exception("WebView2 не инициализирован");

            if (webView.InvokeRequired)
            {
                var tcs = new TaskCompletionSource<object>();
                webView.BeginInvoke(new Action(async () =>
                {
                    try
                    {
                        if (webView.CoreWebView2 == null)
                            throw new Exception("CoreWebView2 не инициализирован");
                        await action(webView.CoreWebView2);
                        tcs.SetResult(null);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }));
                await tcs.Task;
            }
            else
            {
                if (webView.CoreWebView2 == null)
                    throw new Exception("CoreWebView2 не инициализирован");
                await action(webView.CoreWebView2);
            }
        }

        private void InvokeOnUIThread(WebView2 webView, Action<CoreWebView2> action)
        {
            if (webView == null)
                throw new Exception("WebView2 не инициализирован");

            if (webView.InvokeRequired)
            {
                webView.Invoke(new Action(() =>
                {
                    if (webView.CoreWebView2 == null)
                        throw new Exception("CoreWebView2 не инициализирован");
                    action(webView.CoreWebView2);
                }));
            }
            else
            {
                if (webView.CoreWebView2 == null)
                    throw new Exception("CoreWebView2 не инициализирован");
                action(webView.CoreWebView2);
            }
        }

        private string CreateMermaidHtml(string mermaidCode)
        {
            // Экранируем HTML специальные символы в коде Mermaid
            string htmlEscapedCode = mermaidCode
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
            
            return $@"<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"">
    <script src=""https://cdn.jsdelivr.net/npm/mermaid@10.9.0/dist/mermaid.min.js""></script>
    <style>
        body {{
            margin: 0;
            padding: 20px;
            background: white;
            font-family: Arial, sans-serif;
        }}
        .mermaid {{
            background: white;
        }}
    </style>
</head>
<body>
    <div class=""mermaid"" id=""mermaid-diagram"">
{htmlEscapedCode}
    </div>
    <script>
        var isRendered = false;
        
        function checkMermaidRendered() {{
            var diagram = document.getElementById('mermaid-diagram');
            if (!diagram) {{
                return false;
            }}
            
            // Проверяем наличие SVG элементов (Mermaid рендерит в SVG)
            var svg = diagram.querySelector('svg');
            if (svg && svg.children.length > 0) {{
                // Проверяем что SVG имеет размеры
                var width = svg.getAttribute('width') || svg.style.width;
                var height = svg.getAttribute('height') || svg.style.height;
                if (width && height && width !== '0' && height !== '0') {{
                    return true;
                }}
            }}
            
            return false;
        }}
        
        function notifyReady() {{
            if (isRendered) return;
            
            if (checkMermaidRendered()) {{
                isRendered = true;
                if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {{
                    window.chrome.webview.postMessage('mermaidReady');
                }}
            }} else {{
                // Повторная проверка через некоторое время
                setTimeout(function() {{
                    if (checkMermaidRendered()) {{
                        isRendered = true;
                        if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {{
                            window.chrome.webview.postMessage('mermaidReady');
                        }}
                    }} else {{
                        // Если не отрендерилось, все равно отправляем сообщение
                        if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {{
                            window.chrome.webview.postMessage('mermaidReady');
                        }}
                    }}
                }}, 2000);
            }}
        }}
        
        // Инициализируем Mermaid с автоматическим рендерингом
        mermaid.initialize({{
            startOnLoad: true,
            theme: 'default',
            securityLevel: 'loose',
            logLevel: 'error'
        }});
        
        window.addEventListener('load', function() {{
            // Даем время на рендеринг
            setTimeout(notifyReady, 1500);
        }});
        
        // Если DOM уже готов
        if (document.readyState === 'complete' || document.readyState === 'interactive') {{
            setTimeout(notifyReady, 1500);
        }}
    </script>
</body>
</html>";
        }

        private async Task<bool> WaitForMermaidRenderAsync(WebView2 webView)
        {
            var tcs = new TaskCompletionSource<bool>();
            bool messageReceived = false;
            
            void MessageHandler(object sender, CoreWebView2WebMessageReceivedEventArgs e)
            {
                try
                {
                    string message = e.TryGetWebMessageAsString();
                    if (message == "mermaidReady")
                    {
                        messageReceived = true;
                        InvokeOnUIThread(webView, coreWebView =>
                        {
                            coreWebView.WebMessageReceived -= MessageHandler;
                        });
                        tcs.SetResult(true);
                    }
                }
                catch
                {
                }
            }
            
            await InvokeOnUIThreadAsync(webView, async coreWebView =>
            {
                coreWebView.WebMessageReceived += MessageHandler;
            });
            
            await Task.WhenAny(
                tcs.Task,
                Task.Delay(8000)
            );
            
            if (!messageReceived)
            {
                InvokeOnUIThread(webView, coreWebView =>
                {
                    coreWebView.WebMessageReceived -= MessageHandler;
                });
            }
            
            // Дополнительная проверка через JavaScript, что диаграмма действительно отрендерилась
            try
            {
                var checkScript = @"
                    (function() {
                        var diagram = document.getElementById('mermaid-diagram');
                        if (!diagram) return false;
                        var svg = diagram.querySelector('svg');
                        if (svg && svg.children.length > 0) {
                            return true;
                        }
                        var errorDiv = document.getElementById('error-message');
                        if (errorDiv && errorDiv.style.display !== 'none' && errorDiv.textContent.trim() !== '') {
                            return false;
                        }
                        return false;
                    })();
                ";
                
                var result = await InvokeOnUIThreadAsync(webView, async coreWebView =>
                {
                    return await coreWebView.ExecuteScriptAsync(checkScript);
                });
                var isRendered = result?.Trim('"') == "true";
                
                if (!isRendered)
                {
                    await Task.Delay(1000);
                    var result2 = await InvokeOnUIThreadAsync(webView, async coreWebView =>
                    {
                        return await coreWebView.ExecuteScriptAsync(checkScript);
                    });
                    isRendered = result2?.Trim('"') == "true";
                }
                
                return isRendered;
            }
            catch
            {
                await Task.Delay(500);
                return true;
            }
        }

        private async Task<byte[]> CaptureDiagramScreenshotAsync(WebView2 webView)
        {
            try
            {
                // Сначала пробуем получить SVG через JavaScript и конвертировать в PNG
                // Используем подход с postMessage для асинхронной обработки
                var tcs = new TaskCompletionSource<byte[]>();
                byte[] resultData = null;
                bool messageReceived = false;
                
                void MessageHandler(object sender, CoreWebView2WebMessageReceivedEventArgs e)
                {
                    try
                    {
                        string message = e.TryGetWebMessageAsString();
                        if (message != null && message.StartsWith("SVG_TO_PNG:"))
                        {
                            messageReceived = true;
                            var base64Data = message.Substring("SVG_TO_PNG:".Length);
                            if (!string.IsNullOrEmpty(base64Data) && base64Data != "null")
                            {
                                try
                                {
                                    resultData = Convert.FromBase64String(base64Data);
                                    tcs.SetResult(resultData);
                                }
                                catch
                                {
                                    tcs.SetResult(null);
                                }
                            }
                            else
                            {
                                tcs.SetResult(null);
                            }
                            
                            InvokeOnUIThread(webView, coreWebView =>
                            {
                                coreWebView.WebMessageReceived -= MessageHandler;
                            });
                        }
                    }
                    catch
                    {
                        if (!messageReceived)
                        {
                            tcs.SetResult(null);
                        }
                    }
                }
                
                await InvokeOnUIThreadAsync(webView, async coreWebView =>
                {
                    coreWebView.WebMessageReceived += MessageHandler;
                });
                
                var svgToPngScript = @"
                    (function() {
                        try {
                            var diagram = document.getElementById('mermaid-diagram');
                            if (!diagram) {
                                if (window.chrome && window.chrome.webview) {
                                    window.chrome.webview.postMessage('SVG_TO_PNG:null');
                                }
                                return;
                            }
                            
                            var svg = diagram.querySelector('svg');
                            if (!svg) {
                                if (window.chrome && window.chrome.webview) {
                                    window.chrome.webview.postMessage('SVG_TO_PNG:null');
                                }
                                return;
                            }
                            
                            // Получаем размеры SVG
                            var svgWidth = svg.getAttribute('width') || svg.getBoundingClientRect().width || 800;
                            var svgHeight = svg.getAttribute('height') || svg.getBoundingClientRect().height || 600;
                            
                            // Создаем canvas
                            var canvas = document.createElement('canvas');
                            canvas.width = parseInt(svgWidth) || 800;
                            canvas.height = parseInt(svgHeight) || 600;
                            var ctx = canvas.getContext('2d');
                            
                            // Создаем изображение из SVG
                            var svgData = new XMLSerializer().serializeToString(svg);
                            var svgBlob = new Blob([svgData], {type: 'image/svg+xml;charset=utf-8'});
                            var url = URL.createObjectURL(svgBlob);
                            
                            var img = new Image();
                            img.onload = function() {
                                try {
                                    ctx.fillStyle = 'white';
                                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                                    ctx.drawImage(img, 0, 0);
                                    var dataUrl = canvas.toDataURL('image/png');
                                    var base64 = dataUrl.substring('data:image/png;base64,'.length);
                                    URL.revokeObjectURL(url);
                                    if (window.chrome && window.chrome.webview) {
                                        window.chrome.webview.postMessage('SVG_TO_PNG:' + base64);
                                    }
                                } catch(e) {
                                    URL.revokeObjectURL(url);
                                    if (window.chrome && window.chrome.webview) {
                                        window.chrome.webview.postMessage('SVG_TO_PNG:null');
                                    }
                                }
                            };
                            img.onerror = function() {
                                URL.revokeObjectURL(url);
                                if (window.chrome && window.chrome.webview) {
                                    window.chrome.webview.postMessage('SVG_TO_PNG:null');
                                }
                            };
                            img.src = url;
                        } catch(e) {
                            if (window.chrome && window.chrome.webview) {
                                window.chrome.webview.postMessage('SVG_TO_PNG:null');
                            }
                        }
                    })();
                ";
                
                await InvokeOnUIThreadAsync(webView, async coreWebView =>
                {
                    await coreWebView.ExecuteScriptAsync(svgToPngScript);
                });
                
                // Ждем ответа с таймаутом
                var completedTask = await Task.WhenAny(
                    tcs.Task,
                    Task.Delay(5000)
                );
                
                if (completedTask == tcs.Task)
                {
                    var dataUrl = await tcs.Task;
                    if (dataUrl != null && dataUrl.Length > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CaptureDiagramScreenshotAsync] SVG->PNG успешно, размер: {dataUrl.Length} байт");
                        return dataUrl;
                    }
                }
                else
                {
                    // Таймаут - отписываемся от события
                    InvokeOnUIThread(webView, coreWebView =>
                    {
                        coreWebView.WebMessageReceived -= MessageHandler;
                    });
                    System.Diagnostics.Debug.WriteLine("[CaptureDiagramScreenshotAsync] Таймаут при конвертации SVG->PNG");
                }
                
                // Fallback: используем CapturePreviewAsync
                System.Diagnostics.Debug.WriteLine("[CaptureDiagramScreenshotAsync] Используем CapturePreviewAsync как fallback");
                
                // Проверяем размер и видимость WebView2 перед захватом
                bool isVisible = false;
                int width = 0;
                int height = 0;
                
                if (webView.InvokeRequired)
                {
                    webView.Invoke(new Action(() =>
                    {
                        isVisible = webView.Visible && webView.Width > 0 && webView.Height > 0;
                        width = webView.Width;
                        height = webView.Height;
                    }));
                }
                else
                {
                    isVisible = webView.Visible && webView.Width > 0 && webView.Height > 0;
                    width = webView.Width;
                    height = webView.Height;
                }
                
                System.Diagnostics.Debug.WriteLine($"[CaptureDiagramScreenshotAsync] WebView2 видим: {isVisible}, размер: {width}x{height}");
                
                if (!isVisible || width < 100 || height < 100)
                {
                    throw new Exception($"WebView2 не видим или слишком мал: {width}x{height}. Убедитесь, что панель задач Word открыта и имеет достаточный размер.");
                }
                
                return await InvokeOnUIThreadAsync(webView, async coreWebView =>
                {
                    using (var stream = new MemoryStream())
                    {
                        await coreWebView.CapturePreviewAsync(
                            CoreWebView2CapturePreviewImageFormat.Png,
                            stream
                        );
                        
                        var data = stream.ToArray();
                        System.Diagnostics.Debug.WriteLine($"[CaptureDiagramScreenshotAsync] CapturePreviewAsync вернул {data.Length} байт");
                        return data;
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CaptureDiagramScreenshotAsync] Ошибка: {ex.Message}");
                throw new Exception($"Ошибка при создании скриншота: {ex.Message}", ex);
            }
        }

        public async Task<bool> RenderDiagramToPngAsync(
            string mermaidCode, 
            WebView2 webView, 
            string outputPath)
        {
            try
            {
                if (webView == null)
                {
                    System.Diagnostics.Debug.WriteLine("[RenderDiagramToPngAsync] WebView2 не готов для рендеринга");
                    return false;
                }

                // Проверяем размер и видимость WebView2
                bool isVisible = false;
                if (webView.InvokeRequired)
                {
                    webView.Invoke(new Action(() =>
                    {
                        isVisible = webView.Visible && webView.Width > 0 && webView.Height > 0;
                        System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] WebView2 размер: {webView.Width}x{webView.Height}, видим: {webView.Visible}");
                    }));
                }
                else
                {
                    isVisible = webView.Visible && webView.Width > 0 && webView.Height > 0;
                    System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] WebView2 размер: {webView.Width}x{webView.Height}, видим: {webView.Visible}");
                }
                
                if (!isVisible)
                {
                    System.Diagnostics.Debug.WriteLine("[RenderDiagramToPngAsync] Предупреждение: WebView2 может быть невидим");
                }

                string html = CreateMermaidHtml(mermaidCode);
                
                InvokeOnUIThread(webView, coreWebView =>
                {
                    coreWebView.NavigateToString(html);
                });
                
                // Ждем завершения навигации
                await Task.Delay(1500);
                
                // Ждем события NavigationCompleted
                var navigationCompleted = false;
                try
                {
                    await InvokeOnUIThreadAsync(webView, async coreWebView =>
                    {
                        var tcs = new TaskCompletionSource<bool>();
                        void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
                        {
                            coreWebView.NavigationCompleted -= OnNavigationCompleted;
                            tcs.SetResult(true);
                        }
                        coreWebView.NavigationCompleted += OnNavigationCompleted;
                        await tcs.Task;
                        navigationCompleted = true;
                    });
                }
                catch
                {
                    // Если не удалось дождаться события, продолжаем
                    System.Diagnostics.Debug.WriteLine("[RenderDiagramToPngAsync] Не удалось дождаться NavigationCompleted, продолжаем");
                }
                
                System.Diagnostics.Debug.WriteLine("[RenderDiagramToPngAsync] Навигация завершена");
                
                bool isRendered = await WaitForMermaidRenderAsync(webView);
                
                // Дополнительная проверка через JavaScript
                for (int attempt = 0; attempt < 5; attempt++)
                {
                    var checkScript = @"
                        (function() {
                            var diagram = document.getElementById('mermaid-diagram');
                            if (!diagram) return JSON.stringify({rendered: false, reason: 'diagram not found'});
                            var svg = diagram.querySelector('svg');
                            if (svg && svg.children.length > 0) {
                                var width = svg.getAttribute('width') || svg.style.width || svg.getBoundingClientRect().width;
                                var height = svg.getAttribute('height') || svg.style.height || svg.getBoundingClientRect().height;
                                return JSON.stringify({
                                    rendered: true, 
                                    width: width, 
                                    height: height,
                                    childrenCount: svg.children.length
                                });
                            }
                            return JSON.stringify({rendered: false, reason: 'svg not found or empty'});
                        })();
                    ";
                    
                    var result = await InvokeOnUIThreadAsync(webView, async coreWebView =>
                    {
                        return await coreWebView.ExecuteScriptAsync(checkScript);
                    });
                    
                    System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] Проверка рендеринга (попытка {attempt + 1}): {result}");
                    
                    if (!string.IsNullOrEmpty(result) && result.Contains("\"rendered\":true"))
                    {
                        isRendered = true;
                        break;
                    }
                    
                    await Task.Delay(1500);
                }
                
                if (!isRendered)
                {
                    System.Diagnostics.Debug.WriteLine("[RenderDiagramToPngAsync] Предупреждение: диаграмма может быть не полностью отрендерена, но продолжаем");
                    await Task.Delay(2000);
                }
                
                await Task.Delay(1000);
                
                System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] Начинаем создание скриншота для: {outputPath}");
                
                byte[] imageData = await CaptureDiagramScreenshotAsync(webView);
                
                if (imageData == null || imageData.Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[RenderDiagramToPngAsync] Ошибка: данные изображения пусты");
                    return false;
                }
                
                System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] Размер изображения: {imageData.Length} байт");
                
                File.WriteAllBytes(outputPath, imageData);
                
                System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] PNG файл успешно сохранен: {outputPath}");
                
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[RenderDiagramToPngAsync] Ошибка рендеринга диаграммы: {ex.Message}\n{ex.StackTrace}");
                return false;
            }
        }

        public async Task<ExportResult> ExportAllDiagramsToPngAsync(
            string markdown,
            WebView2 webView,
            string outputFolder,
            IProgress<ExportProgress> progress,
            CancellationToken cancellationToken)
        {
            var result = new ExportResult
            {
                TotalDiagrams = 0,
                SuccessCount = 0,
                FailedCount = 0,
                Errors = new List<string>()
            };
            
            try
            {
                var diagrams = ExtractMermaidDiagrams(markdown);
                result.TotalDiagrams = diagrams.Count;
                
                if (diagrams.Count == 0)
                {
                    result.Errors.Add("Диаграммы Mermaid не найдены в markdown");
                    return result;
                }
                
                for (int i = 0; i < diagrams.Count; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        result.Errors.Add("Операция отменена пользователем");
                        break;
                    }
                    
                    string fileName = $"mermaid_{i + 1}.png";
                    string outputPath = Path.Combine(outputFolder, fileName);
                    
                    progress?.Report(new ExportProgress
                    {
                        CurrentIndex = i + 1,
                        TotalCount = diagrams.Count,
                        CurrentFileName = fileName,
                        Status = "Обработка..."
                    });
                    
                    System.Diagnostics.Debug.WriteLine($"[ExportAllDiagramsToPngAsync] Обработка диаграммы {i + 1} из {diagrams.Count}: {fileName}");
                    
                    bool success = await RenderDiagramToPngAsync(
                        diagrams[i], 
                        webView, 
                        outputPath
                    );
                    
                    if (success)
                    {
                        result.SuccessCount++;
                        System.Diagnostics.Debug.WriteLine($"[ExportAllDiagramsToPngAsync] Диаграмма {i + 1} успешно экспортирована");
                    }
                    else
                    {
                        result.FailedCount++;
                        string errorMsg = $"Не удалось обработать диаграмму {i + 1}: {fileName}";
                        result.Errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[ExportAllDiagramsToPngAsync] Ошибка: {errorMsg}");
                    }
                }
                
                return result;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Критическая ошибка: {ex.Message}");
                return result;
            }
        }
    }

    public class ExportResult
    {
        public int TotalDiagrams { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    public class ExportProgress
    {
        public int CurrentIndex { get; set; }
        public int TotalCount { get; set; }
        public string CurrentFileName { get; set; }
        public string Status { get; set; }
    }
}
