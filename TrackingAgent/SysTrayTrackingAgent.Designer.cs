
namespace TrackingAgent
{
    partial class SysTrayTrackingAgent
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
                trayIcon.Dispose();
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
            this.components = new System.ComponentModel.Container();
            this.myNotifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.status = new System.Windows.Forms.Label();
            this.close = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // myNotifyIcon
            // 
            this.myNotifyIcon.Text = "NotifyIcon";
            this.myNotifyIcon.Visible = true;
            // 
            // status
            // 
            this.status.AutoSize = true;
            this.status.Location = new System.Drawing.Point(36, 19);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(53, 13);
            this.status.TabIndex = 0;
            this.status.Text = "Running..";
            // 
            // close
            // 
            this.close.Location = new System.Drawing.Point(130, 14);
            this.close.Name = "close";
            this.close.Size = new System.Drawing.Size(75, 23);
            this.close.TabIndex = 1;
            this.close.Text = "Close";
            this.close.UseVisualStyleBackColor = true;
            this.close.Click += new System.EventHandler(this.close_Click);
            // 
            // SysTrayTrackingAgent
            // 
            this.ClientSize = new System.Drawing.Size(226, 52);
            this.ControlBox = false;
            this.Controls.Add(this.close);
            this.Controls.Add(this.status);
            this.Name = "SysTrayTrackingAgent";
            this.Text = "Tracking Agent";
            this.Load += new System.EventHandler(this.SysTrayTrackingAgent_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.NotifyIcon myNotifyIcon;
        private System.Windows.Forms.Label status;
        private System.Windows.Forms.Button close;
    }
}

