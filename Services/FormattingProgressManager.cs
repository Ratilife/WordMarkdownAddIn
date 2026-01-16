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
        private System.Windows.Forms.Timer _showTimer;
        private string _currentOperationName;
        private readonly object _lockObject = new object();
        private bool _disposed = false;

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
                if (_disposed)
                    return;

                _currentOperationName = operationName;
                _stopwatch.Restart();
                _isProgressVisible = false;

                // Запускаем таймер для показа окна через threshold секунд
                // Используем Windows.Forms.Timer, который работает в UI потоке (STA)
                _showTimer?.Stop();
                _showTimer?.Dispose();
                _showTimer = new System.Windows.Forms.Timer();
                _showTimer.Interval = (int)_threshold.TotalMilliseconds;
                _showTimer.Tick += (s, e) =>
                {
                    _showTimer.Stop();
                    OnShowTimerElapsed(null);
                };
                _showTimer.Start();
            }
        }

        private void OnShowTimerElapsed(object state)
        {
            lock (_lockObject)
            {
                if (_disposed || _isProgressVisible || _stopwatch == null || !_stopwatch.IsRunning)
                    return;

                ShowProgressWindow(_currentOperationName);
            }
        }

        private void ShowProgressWindow(string operationName)
        {
            try
            {
                // В VSTO все операции выполняются в UI потоке Word (STA)
                // Создаем окно напрямую, так как мы уже в правильном потоке
                if (_progressWindow == null || _progressWindow.IsDisposed)
                {
                    _progressWindow = new ProgressWindow();
                }
                _progressWindow.SetOperationName(operationName);
                _progressWindow.ProgressValue = 0;
                _progressWindow.CurrentStage = "Начало обработки...";
                _progressWindow.Show();
                _isProgressVisible = true;
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FormattingProgressManager] Ошибка при показе окна: {ex.Message}");
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
                if (_disposed)
                    return;

                // Если окно еще не показано, но прошло достаточно времени - показываем
                if (!_isProgressVisible && _stopwatch != null && _stopwatch.IsRunning && _stopwatch.Elapsed >= _threshold)
                {
                    ShowProgressWindow(_currentOperationName);
                }

                if (_isProgressVisible && _progressWindow != null && !_progressWindow.IsDisposed)
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
                if (_disposed)
                    return;

                // Останавливаем таймер
                _showTimer?.Dispose();
                _showTimer = null;

                // Останавливаем секундомер
                _stopwatch?.Stop();

                // Закрываем окно прогресса
                if (_isProgressVisible && _progressWindow != null && !_progressWindow.IsDisposed)
                {
                    try
                    {
                        // В VSTO мы уже в UI потоке, можно закрывать напрямую
                        _progressWindow.Hide();
                        _progressWindow.Close();
                        _progressWindow.Dispose();
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
        /// Отмена операции (для будущих версий)
        /// </summary>
        public void CancelOperation()
        {
            CompleteOperation();
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                CompleteOperation();
                _stopwatch?.Stop();
                _stopwatch = null;
                _disposed = true;
            }
        }
    }
}
