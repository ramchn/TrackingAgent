﻿
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SysTrayTrackingAgent));
            this.myNotifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.SuspendLayout();
            // 
            // myNotifyIcon
            // 
            this.myNotifyIcon.Text = "NotifyIcon";
            this.myNotifyIcon.Visible = true;
            // 
            // SysTrayTrackingAgent
            // 
            this.ClientSize = new System.Drawing.Size(226, 52);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SysTrayTrackingAgent";
            this.Text = "Tracking Agent";
            this.Load += new System.EventHandler(this.SysTrayTrackingAgent_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.NotifyIcon myNotifyIcon;
    }
}

