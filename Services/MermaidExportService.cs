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

        private string CreateMermaidHtml(string mermaidCode)
        {
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
    <div class=""mermaid"">
{mermaidCode}
    </div>
    <script>
        mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
        
        window.addEventListener('load', function() {{
            setTimeout(function() {{
                if (window.chrome && window.chrome.webview) {{
                    window.chrome.webview.postMessage('mermaidReady');
                }}
            }}, 1000);
        }});
    </script>
</body>
</html>";
        }

        private async Task WaitForMermaidRenderAsync(WebView2 webView)
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
                        webView.CoreWebView2.WebMessageReceived -= MessageHandler;
                        tcs.SetResult(true);
                    }
                }
                catch
                {
                }
            }
            
            if (webView?.CoreWebView2 != null)
            {
                webView.CoreWebView2.WebMessageReceived += MessageHandler;
            }
            
            await Task.WhenAny(
                tcs.Task,
                Task.Delay(5000)
            );
            
            if (!messageReceived && webView?.CoreWebView2 != null)
            {
                webView.CoreWebView2.WebMessageReceived -= MessageHandler;
            }
            
            await Task.Delay(500);
        }

        private async Task<byte[]> CaptureDiagramScreenshotAsync(WebView2 webView)
        {
            try
            {
                if (webView?.CoreWebView2 == null)
                {
                    throw new Exception("WebView2 не инициализирован");
                }

                using (var stream = new MemoryStream())
                {
                    await webView.CoreWebView2.CapturePreviewAsync(
                        CoreWebView2CapturePreviewImageFormat.Png,
                        stream
                    );
                    
                    return stream.ToArray();
                }
            }
            catch (Exception ex)
            {
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
                if (webView?.CoreWebView2 == null)
                {
                    System.Diagnostics.Debug.WriteLine("WebView2 не готов для рендеринга");
                    return false;
                }

                string html = CreateMermaidHtml(mermaidCode);
                
                webView.CoreWebView2.NavigateToString(html);
                
                await Task.Delay(500);
                
                await WaitForMermaidRenderAsync(webView);
                
                byte[] imageData = await CaptureDiagramScreenshotAsync(webView);
                
                File.WriteAllBytes(outputPath, imageData);
                
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка рендеринга диаграммы: {ex.Message}");
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
                    
                    bool success = await RenderDiagramToPngAsync(
                        diagrams[i], 
                        webView, 
                        outputPath
                    );
                    
                    if (success)
                    {
                        result.SuccessCount++;
                    }
                    else
                    {
                        result.FailedCount++;
                        result.Errors.Add($"Не удалось обработать диаграмму {i + 1}: {fileName}");
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
}
