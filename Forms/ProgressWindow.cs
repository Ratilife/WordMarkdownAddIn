using System;
using System.Windows.Forms;

namespace WordMarkdownAddIn.Forms
{
    /// <summary>
    /// Форма для отображения прогресса операции форматирования
    /// </summary>
    public partial class ProgressWindow : Form
    {
        private string _operationName;
        private string _currentStage;
        private int _progressValue;

        /// <summary>
        /// Название операции
        /// </summary>
        public string OperationName
        {
            get => _operationName;
            set
            {
                _operationName = value;
                if (lblOperation != null)
                {
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => lblOperation.Text = value));
                    }
                    else
                    {
                        lblOperation.Text = value;
                    }
                }
            }
        }

        /// <summary>
        /// Текущий этап обработки
        /// </summary>
        public string CurrentStage
        {
            get => _currentStage;
            set
            {
                _currentStage = value;
                if (lblStage != null)
                {
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => lblStage.Text = value));
                    }
                    else
                    {
                        lblStage.Text = value;
                    }
                }
            }
        }

        /// <summary>
        /// Значение прогресса (0-100)
        /// </summary>
        public int ProgressValue
        {
            get => _progressValue;
            set
            {
                _progressValue = Math.Max(0, Math.Min(100, value));
                if (progressBar != null)
                {
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => progressBar.Value = _progressValue));
                    }
                    else
                    {
                        progressBar.Value = _progressValue;
                    }
                }
            }
        }

        public ProgressWindow()
        {
            InitializeComponent();
            _operationName = string.Empty;
            _currentStage = string.Empty;
            _progressValue = 0;
        }

        /// <summary>
        /// Установить название операции
        /// </summary>
        public void SetOperationName(string operationName)
        {
            OperationName = operationName;
        }

        /// <summary>
        /// Обновить прогресс и текст этапа
        /// </summary>
        public void UpdateProgress(int value, string stage)
        {
            ProgressValue = value;
            CurrentStage = stage;
            Application.DoEvents();
        }

        /// <summary>
        /// Показать окно прогресса
        /// </summary>
        public void ShowProgress()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Show()));
            }
            else
            {
                Show();
            }
        }

        /// <summary>
        /// Скрыть окно прогресса
        /// </summary>
        public void HideProgress()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Hide()));
            }
            else
            {
                Hide();
            }
        }
    }
}
