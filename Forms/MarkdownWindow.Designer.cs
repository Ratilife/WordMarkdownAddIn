namespace WordMarkdownAddIn.Forms
{
    partial class MarkdownWindow
    {
        private System.ComponentModel.IContainer components = null;
        private WordMarkdownAddIn.Controls.TaskPaneControl taskPaneControl;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.taskPaneControl = new WordMarkdownAddIn.Controls.TaskPaneControl();
            this.SuspendLayout();
            
            // taskPaneControl
            this.taskPaneControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.taskPaneControl.Location = new System.Drawing.Point(0, 0);
            this.taskPaneControl.Name = "taskPaneControl";
            this.taskPaneControl.Size = new System.Drawing.Size(800, 600);
            this.taskPaneControl.TabIndex = 0;
            
            // MarkdownWindow
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 600);
            this.Controls.Add(this.taskPaneControl);
            this.MinimumSize = new System.Drawing.Size(400, 300);
            this.Name = "MarkdownWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Markdown Editor";
            this.WindowState = System.Windows.Forms.FormWindowState.Normal;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.ResumeLayout(false);
        }
    }
}
