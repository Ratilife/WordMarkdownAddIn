using System;
using System.Drawing;
using System.Windows.Forms;
using WordMarkdownAddIn.Controls;

namespace WordMarkdownAddIn.Forms
{
    public partial class MarkdownWindow : Form
    {
        public MarkdownWindow()
        {
            InitializeComponent();
            SetupWindow();
        }

        private void SetupWindow()
        {
            // Устанавливаем начальный размер - 80% экрана
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Width = (int)(screen.Width * 0.8);
            this.Height = (int)(screen.Height * 0.8);
            
            // Центрируем окно
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        public TaskPaneControl GetTaskPaneControl()
        {
            return taskPaneControl;
        }

        public void SetMarkdown(string markdown)
        {
            if (taskPaneControl != null)
            {
                taskPaneControl.SetMarkdown(markdown);
            }
        }

        public string GetMarkdown()
        {
            if (taskPaneControl != null)
            {
                return taskPaneControl.GetCachedMarkdown();
            }
            return string.Empty;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Сохраняем размер и позицию окна
            if (ThisAddIn.Instance != null)
            {
                ThisAddIn.Instance.SaveMarkdownWindowState(this);
                
                // При закрытии окна показываем панель обратно
                var pane = ThisAddIn.MarkdownPane;
                if (pane != null)
                {
                    // Синхронизируем содержимое обратно в панель
                    var paneControl = ThisAddIn.PaneControl;
                    if (paneControl != null)
                    {
                        string markdown = this.GetMarkdown();
                        if (!string.IsNullOrEmpty(markdown))
                        {
                            paneControl.SetMarkdown(markdown);
                        }
                    }
                    pane.Visible = true;
                }
                
                // Обнуляем ссылку на окно в ThisAddIn
                ThisAddIn.ClearMarkdownWindow();
            }
            base.OnFormClosing(e);
        }
    }
}
