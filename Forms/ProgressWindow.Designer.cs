namespace WordMarkdownAddIn.Forms
{
    partial class ProgressWindow
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblOperation;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStage;

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
            this.lblOperation = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStage = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // lblOperation
            this.lblOperation.AutoSize = true;
            this.lblOperation.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblOperation.Location = new System.Drawing.Point(12, 15);
            this.lblOperation.Name = "lblOperation";
            this.lblOperation.Size = new System.Drawing.Size(0, 15);
            this.lblOperation.TabIndex = 0;

            // progressBar
            this.progressBar.Location = new System.Drawing.Point(15, 45);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(370, 23);
            this.progressBar.TabIndex = 1;
            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = 100;
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;

            // lblStage
            this.lblStage.AutoSize = true;
            this.lblStage.Location = new System.Drawing.Point(12, 85);
            this.lblStage.Name = "lblStage";
            this.lblStage.Size = new System.Drawing.Size(0, 13);
            this.lblStage.TabIndex = 2;

            // ProgressWindow
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 150);
            this.Controls.Add(this.lblStage);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblOperation);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Обработка";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
