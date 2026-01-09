using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    }
}
