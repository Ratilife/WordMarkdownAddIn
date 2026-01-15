using System;
using System.Windows.Forms;

namespace WordMarkdownAddIn
{
    public partial class ProgressForm : Form
    {
        public bool IsCancelled { get; private set; } = false;
        
        public ProgressForm()
        {
            InitializeComponent();
        }
        
        public void UpdateProgress(int current, int total, string fileName)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateProgress(current, total, fileName)));
                return;
            }
            
            int percentage = total > 0 ? (current * 100) / total : 0;
            progressBar.Value = Math.Min(percentage, 100);
            
            lblStatus.Text = $"Обработка {current} из {total}: {fileName}";
            Application.DoEvents();
        }
        
        private void btnCancel_Click(object sender, EventArgs e)
        {
            IsCancelled = true;
            btnCancel.Enabled = false;
            lblStatus.Text = "Отмена...";
        }
    }
}
