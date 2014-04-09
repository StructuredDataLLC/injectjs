namespace InjectExcel
{
    partial class scriptPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tbEditor = new System.Windows.Forms.TextBox();
            this.tbLog = new System.Windows.Forms.TextBox();
            this.bExecute = new System.Windows.Forms.Button();
            this.bClearLog = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.bClearLog);
            this.splitContainer1.Panel1.Controls.Add(this.bExecute);
            this.splitContainer1.Panel1.Controls.Add(this.tbEditor);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tbLog);
            this.splitContainer1.Size = new System.Drawing.Size(629, 593);
            this.splitContainer1.SplitterDistance = 466;
            this.splitContainer1.TabIndex = 0;
            // 
            // tbEditor
            // 
            this.tbEditor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbEditor.Location = new System.Drawing.Point(3, 34);
            this.tbEditor.Multiline = true;
            this.tbEditor.Name = "tbEditor";
            this.tbEditor.Size = new System.Drawing.Size(623, 400);
            this.tbEditor.TabIndex = 0;
            // 
            // tbLog
            // 
            this.tbLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbLog.Location = new System.Drawing.Point(3, 3);
            this.tbLog.Multiline = true;
            this.tbLog.Name = "tbLog";
            this.tbLog.Size = new System.Drawing.Size(623, 117);
            this.tbLog.TabIndex = 1;
            // 
            // bExecute
            // 
            this.bExecute.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bExecute.Location = new System.Drawing.Point(3, 440);
            this.bExecute.Name = "bExecute";
            this.bExecute.Size = new System.Drawing.Size(75, 23);
            this.bExecute.TabIndex = 1;
            this.bExecute.Text = "Execute";
            this.bExecute.UseVisualStyleBackColor = true;
            // 
            // bClearLog
            // 
            this.bClearLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bClearLog.Location = new System.Drawing.Point(84, 440);
            this.bClearLog.Name = "bClearLog";
            this.bClearLog.Size = new System.Drawing.Size(75, 23);
            this.bClearLog.TabIndex = 2;
            this.bClearLog.Text = "Clear Log";
            this.bClearLog.UseVisualStyleBackColor = true;
            // 
            // scriptPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "scriptPane";
            this.Size = new System.Drawing.Size(629, 593);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button bClearLog;
        private System.Windows.Forms.Button bExecute;
        private System.Windows.Forms.TextBox tbEditor;
        private System.Windows.Forms.TextBox tbLog;
    }
}
