using System.Drawing;

namespace Markdown4Outlook
{
    partial class MyTaskPane
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.SplitContainer splitContainer;
        private System.Windows.Forms.TextBox inputBox;
        private System.Windows.Forms.WebBrowser previewBrowser;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        private void InitializeComponent()
        {
        	this.splitContainer = new System.Windows.Forms.SplitContainer();
        	this.inputBox = new System.Windows.Forms.TextBox();
        	this.previewBrowser = new System.Windows.Forms.WebBrowser();
        	((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
        	this.splitContainer.Panel1.SuspendLayout();
        	this.splitContainer.Panel2.SuspendLayout();
        	this.splitContainer.SuspendLayout();
        	this.SuspendLayout();
        	// 
        	// splitContainer
        	// 
        	this.splitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
        	this.splitContainer.Location = new System.Drawing.Point(0, 0);
        	this.splitContainer.Name = "splitContainer1";
        	// 
        	// splitContainer1.Panel1
        	// 
        	this.splitContainer.Panel1.Controls.Add(this.inputBox);
        	this.splitContainer.Panel1MinSize = 50;
        	// 
        	// splitContainer1.Panel2
        	// 
        	this.splitContainer.Panel2.Controls.Add(this.previewBrowser);
        	this.splitContainer.Size = new System.Drawing.Size(747, 607);
        	this.splitContainer.SplitterDistance = 321;
        	this.splitContainer.TabIndex = 0;
        	// 
        	// textBox1
        	// 
        	this.inputBox.Dock = System.Windows.Forms.DockStyle.Fill;
        	this.inputBox.Location = new System.Drawing.Point(0, 0);
        	this.inputBox.Multiline = true;
        	this.inputBox.Name = "inputBox";
        	this.inputBox.Size = new System.Drawing.Size(321, 607);
        	this.inputBox.TabIndex = 0;
        	this.inputBox.Font = SystemFonts.MessageBoxFont;
        	this.inputBox.Text = Constants.INIT_TEXT;
        	this.inputBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.UpateHTMLPreview);
        	
        	// 
        	// webBrowser1
        	// 
        	this.previewBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
        	this.previewBrowser.Location = new System.Drawing.Point(0, 0);
        	this.previewBrowser.MinimumSize = new System.Drawing.Size(50, 20);
        	this.previewBrowser.Name = "webBrowser1";
        	this.previewBrowser.Size = new System.Drawing.Size(422, 607);
        	this.previewBrowser.TabIndex = 0;
        	// 
        	// MyTaskPane
        	// 
        	this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        	this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        	this.Controls.Add(this.splitContainer);
        	this.Name = Constants.TASK_PANE_TITLE;
        	this.Size = new System.Drawing.Size(747, 607);
        	this.splitContainer.Panel1.ResumeLayout(false);
        	this.splitContainer.Panel1.PerformLayout();
        	this.splitContainer.Panel2.ResumeLayout(false);
        	((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
        	this.splitContainer.ResumeLayout(false);
        	this.ResumeLayout(false);
        }

        #endregion
    }
}
