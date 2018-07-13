namespace ExcelIdea
{
    partial class StatusForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.barProgress = new System.Windows.Forms.ProgressBar();
            this.txtConsoleLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // barProgress
            // 
            this.barProgress.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.barProgress.Location = new System.Drawing.Point(12, 12);
            this.barProgress.MarqueeAnimationSpeed = 80;
            this.barProgress.Name = "barProgress";
            this.barProgress.Size = new System.Drawing.Size(430, 24);
            this.barProgress.Step = 7;
            this.barProgress.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.barProgress.TabIndex = 0;
            // 
            // txtConsoleLog
            // 
            this.txtConsoleLog.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtConsoleLog.Location = new System.Drawing.Point(12, 42);
            this.txtConsoleLog.Multiline = true;
            this.txtConsoleLog.Name = "txtConsoleLog";
            this.txtConsoleLog.ReadOnly = true;
            this.txtConsoleLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtConsoleLog.Size = new System.Drawing.Size(430, 152);
            this.txtConsoleLog.TabIndex = 1;
            this.txtConsoleLog.TabStop = false;
            // 
            // StatusForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(454, 204);
            this.ControlBox = false;
            this.Controls.Add(this.txtConsoleLog);
            this.Controls.Add(this.barProgress);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "StatusForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelIdea Calculating...";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar barProgress;
        private System.Windows.Forms.TextBox txtConsoleLog;
    }
}