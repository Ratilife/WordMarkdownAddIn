using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Tasks = System.Threading.Tasks;
using System.Windows.Forms;

namespace WordMarkdownAddIn
{
    public partial class MarkdownRibbon
    {
        private void MarkdownRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.SaveMarkdownFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void btnPanel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.TogglePane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void btnOpen_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.OpenMarkdownFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bBold_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertInline("**", "**");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bItalic_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertInline("*", "*");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bStrike_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertInline("~~", "~~");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bH1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertHeading(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }

        }

        private void bH2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertHeading(2);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bH3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertHeading(3);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertBulletList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bNumList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertNumberedList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bCheckbox_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertCheckbox(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bCheckboxTrue_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertCheckbox(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertTable(3, 3);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }

        }

        private void bLink_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertLink("текст", "https://example.com");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bHR_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertSnippet("\n\n---\n\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }

        }

        private void bMermaid_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertMermaidSample();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }

        }


        private void bCodeBlock_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl?.InsertCodeBlock("csharp");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void bMath_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.PaneControl.InsertMathSample();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var service = new Services.WordToMarkdownService();
                string markdown = service.ConvertToMarkdown();
                ThisAddIn.PaneControl.SetMarkdown(markdown);
                MessageBox.Show("Документ Word успешно преобразован в Markdown!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при преобразовании: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnApplyToWord_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 1. Получить Markdown из редактора
                string markdown = ThisAddIn.PaneControl.GetCachedMarkdown();

                // 2. Преобразовать и вставить в Word
                var formatter = new Services.MarkdownToWordFormatter();
                formatter.ApplyMarkdownToWord(markdown);

                MessageBox.Show("Markdown успешно применен в Word!", "Успех");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private void btnConvertMD_DocNotF_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Создаем экземпляр сервиса для извлечения текста
                var service = new Services.WordToMarkdownPlainTextService();

                // Извлекаем текстовое содержимое документа
                string markdown = service.ExtractPlainText();

                // Устанавливаем извлеченный текст в панель Markdown
                ThisAddIn.PaneControl.SetMarkdown(markdown);

                // Показываем сообщение об успешном выполнении
                MessageBox.Show(
                    "Текст из документа Word успешно перенесен в Markdown!",
                    "Успех",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                // Показываем сообщение об ошибке
                MessageBox.Show(
                    $"Ошибка при переносе текста: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }

        }

        private void btnFormatMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var formatter = new Services.WordMarkdownFormatter();

                // Проверяем, есть ли выделенный текст
                Range selection = Globals.ThisAddIn?.Application?.Selection?.Range;
                if (selection != null && !string.IsNullOrEmpty(selection.Text))
                {
                    // Форматируем только выделенный текст
                    formatter.FormatSelectedText();
                    MessageBox.Show(
                        "Markdown-синтаксис в выделенном тексте успешно отформатирован!",
                        "Успех",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    // Форматируем весь документ
                    var result = MessageBox.Show(
                        "Текст не выделен. Отформатировать весь документ?",
                        "Подтверждение",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (result == DialogResult.Yes)
                    {
                        formatter.FormatEntireDocument();
                        MessageBox.Show(
                            "Markdown-синтаксис в документе успешно отформатирован!",
                            "Успех",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при форматировании Markdown: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void bImage_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private async void btnExportMermaid_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (ThisAddIn.PaneControl == null)
                {
                    MessageBox.Show(
                        "Панель управления не инициализирована.",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return;
                }
                
                string markdown = ThisAddIn.PaneControl.GetCachedMarkdown();
                
                if (string.IsNullOrEmpty(markdown))
                {
                    MessageBox.Show(
                        "Панель markdown пуста. Нет данных для экспорта.",
                        "Информация",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }
                
                var webView = ThisAddIn.PaneControl.GetWebView();
                if (webView == null || webView.CoreWebView2 == null)
                {
                    MessageBox.Show(
                        "WebView2 не готов. Подождите загрузки панели.",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return;
                }
                
                // Переключаем панель в HTML режим перед экспортом
                ThisAddIn.PaneControl.SetViewMode("html");
                await Tasks.Task.Delay(300); // Даем время на переключение режима
                
                // Убеждаемся, что кнопки видимы
                if (webView != null)
                {
                    var script = @"
                        (function() {
                            var viewControls = document.querySelector('.view-controls');
                            if (viewControls) {
                                viewControls.style.display = 'flex';
                                viewControls.style.visibility = 'visible';
                            }
                            var btnSplit = document.getElementById('btn-split');
                            var btnMarkdown = document.getElementById('btn-markdown');
                            var btnHtml = document.getElementById('btn-html');
                            if (btnSplit) btnSplit.style.display = 'block';
                            if (btnMarkdown) btnMarkdown.style.display = 'block';
                            if (btnHtml) btnHtml.style.display = 'block';
                        })();
                    ";
                    
                    if (webView.InvokeRequired)
                    {
                        var tcs = new Tasks.TaskCompletionSource<object>();
                        webView.BeginInvoke(new Action(async () =>
                        {
                            try
                            {
                                if (webView.CoreWebView2 != null)
                                {
                                    await webView.CoreWebView2.ExecuteScriptAsync(script);
                                }
                                tcs.SetResult(null);
                            }
                            catch (Exception ex)
                            {
                                tcs.SetException(ex);
                            }
                        }));
                        await tcs.Task;
                    }
                    else if (webView.CoreWebView2 != null)
                    {
                        await webView.CoreWebView2.ExecuteScriptAsync(script);
                    }
                }
                
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Выберите папку для сохранения PNG файлов";
                    folderDialog.ShowNewFolderButton = true;
                    
                    if (folderDialog.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }
                    
                    string outputFolder = folderDialog.SelectedPath;
                    
                    var exportService = new Services.MermaidExportService();
                    
                    var diagrams = exportService.ExtractMermaidDiagrams(markdown);
                    if (diagrams.Count == 0)
                    {
                        MessageBox.Show(
                            "В markdown не найдено диаграмм Mermaid.\n\nДиаграммы должны быть в формате:\n```mermaid\n...\n```",
                            "Информация",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }
                    
                    using (var progressForm = new ProgressForm())
                    {
                        var cancellationTokenSource = new CancellationTokenSource();
                        
                        progressForm.FormClosing += (s, args) =>
                        {
                            if (progressForm.IsCancelled)
                            {
                                cancellationTokenSource.Cancel();
                            }
                        };
                        
                        progressForm.Show();
                        
                        var progress = new Progress<Services.ExportProgress>(p =>
                        {
                            if (!progressForm.IsCancelled)
                            {
                                progressForm.UpdateProgress(p.CurrentIndex, p.TotalCount, p.CurrentFileName);
                            }
                        });
                        
                        var result = await exportService.ExportAllDiagramsToPngAsync(
                            markdown,
                            webView,
                            outputFolder,
                            progress,
                            cancellationTokenSource.Token
                        );
                        
                        progressForm.Close();
                        
                        // Восстанавливаем панель в HTML режиме с кнопками и диаграммой
                        // RestoreHtmlShellAsync уже переключает панель в HTML режим
                        if (ThisAddIn.PaneControl != null)
                        {
                            await ThisAddIn.PaneControl.RestoreHtmlShellAsync(markdown);
                        }
                        
                        string message = $"Экспорт завершен!\n\n" +
                                       $"Всего диаграмм: {result.TotalDiagrams}\n" +
                                       $"Успешно: {result.SuccessCount}\n" +
                                       $"Ошибок: {result.FailedCount}";
                        
                        if (result.Errors.Count > 0)
                        {
                            message += $"\n\nОшибки:\n{string.Join("\n", result.Errors)}";
                        }
                        
                        MessageBox.Show(
                            message,
                            result.FailedCount == 0 ? "Успех" : "Завершено с ошибками",
                            MessageBoxButtons.OK,
                            result.FailedCount == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                // Восстанавливаем HTML оболочку даже при ошибке
                try
                {
                    if (ThisAddIn.PaneControl != null)
                    {
                        string markdown = ThisAddIn.PaneControl.GetCachedMarkdown();
                        await ThisAddIn.PaneControl.RestoreHtmlShellAsync(markdown);
                        
                        // Убеждаемся, что панель в HTML режиме
                        ThisAddIn.PaneControl.SetViewMode("html");
                    }
                }
                catch
                {
                    // Игнорируем ошибки восстановления
                }
                
                MessageBox.Show(
                    $"Ошибка при экспорте: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}
