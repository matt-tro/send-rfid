namespace SendRFID
{
    partial class FormSendKeys
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSendKeys));
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.notifyMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripEsci = new System.Windows.Forms.ToolStripMenuItem();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.notifyMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIcon
            // 
            this.notifyIcon.ContextMenuStrip = this.notifyMenu;
            this.notifyIcon.Text = "SendRFID";
            this.notifyIcon.Visible = true;
            // 
            // notifyMenu
            // 
            this.notifyMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripEsci});
            this.notifyMenu.Name = "notifyMenu";
            this.notifyMenu.Size = new System.Drawing.Size(95, 26);
            // 
            // toolStripEsci
            // 
            this.toolStripEsci.Name = "toolStripEsci";
            this.toolStripEsci.Size = new System.Drawing.Size(94, 22);
            this.toolStripEsci.Text = "&Esci";
            this.toolStripEsci.Click += new System.EventHandler(this.toolStripEsci_Click);
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerSupportsCancellation = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            // 
            // FormSendKeys
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "FormSendKeys";
            this.Text = "Form1";
            this.Shown += new System.EventHandler(this.FormSendKeys_Shown);
            this.notifyMenu.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.ContextMenuStrip notifyMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripEsci;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
    }
}

