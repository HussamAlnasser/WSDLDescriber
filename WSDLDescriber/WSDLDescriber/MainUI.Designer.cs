namespace WSDLDescriber
{
    partial class MainUI
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
            this.components = new System.ComponentModel.Container();
            this.WSDLURLLabel = new System.Windows.Forms.Label();
            this.urlBox = new System.Windows.Forms.TextBox();
            this.AuthorNameLabel = new System.Windows.Forms.Label();
            this.authorNameBox = new System.Windows.Forms.TextBox();
            this.generatorButton = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.statusLabel = new System.Windows.Forms.Label();
            this.statusBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // WSDLURLLabel
            // 
            this.WSDLURLLabel.AutoSize = true;
            this.WSDLURLLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.WSDLURLLabel.Location = new System.Drawing.Point(13, 13);
            this.WSDLURLLabel.Name = "WSDLURLLabel";
            this.WSDLURLLabel.Size = new System.Drawing.Size(72, 13);
            this.WSDLURLLabel.TabIndex = 0;
            this.WSDLURLLabel.Text = "WSDL URL";
            // 
            // urlBox
            // 
            this.urlBox.BackColor = System.Drawing.SystemColors.Window;
            this.urlBox.Location = new System.Drawing.Point(13, 30);
            this.urlBox.Name = "urlBox";
            this.urlBox.Size = new System.Drawing.Size(459, 20);
            this.urlBox.TabIndex = 1;
            // 
            // AuthorNameLabel
            // 
            this.AuthorNameLabel.AutoSize = true;
            this.AuthorNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AuthorNameLabel.Location = new System.Drawing.Point(13, 60);
            this.AuthorNameLabel.Name = "AuthorNameLabel";
            this.AuthorNameLabel.Size = new System.Drawing.Size(80, 13);
            this.AuthorNameLabel.TabIndex = 2;
            this.AuthorNameLabel.Text = "Author Name";
            // 
            // authorNameBox
            // 
            this.authorNameBox.Location = new System.Drawing.Point(99, 57);
            this.authorNameBox.Name = "authorNameBox";
            this.authorNameBox.Size = new System.Drawing.Size(100, 20);
            this.authorNameBox.TabIndex = 3;
            // 
            // generatorButton
            // 
            this.generatorButton.Location = new System.Drawing.Point(12, 227);
            this.generatorButton.Name = "generatorButton";
            this.generatorButton.Size = new System.Drawing.Size(152, 23);
            this.generatorButton.TabIndex = 4;
            this.generatorButton.Text = "Generate Word Document";
            this.generatorButton.UseVisualStyleBackColor = true;
            this.generatorButton.Click += new System.EventHandler(this.generatorButton_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 256);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(460, 20);
            this.progressBar.TabIndex = 5;
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog_FileOk);
            // 
            // timer
            // 
            this.timer.Tick += new System.EventHandler(this.timer_Tick);
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusLabel.Location = new System.Drawing.Point(13, 89);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(43, 13);
            this.statusLabel.TabIndex = 6;
            this.statusLabel.Text = "Status";
            // 
            // statusBox
            // 
            this.statusBox.Enabled = false;
            this.statusBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusBox.Location = new System.Drawing.Point(16, 106);
            this.statusBox.Name = "statusBox";
            this.statusBox.Size = new System.Drawing.Size(456, 115);
            this.statusBox.TabIndex = 7;
            this.statusBox.Text = "";
            // 
            // MainUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 281);
            this.Controls.Add(this.statusBox);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.generatorButton);
            this.Controls.Add(this.authorNameBox);
            this.Controls.Add(this.AuthorNameLabel);
            this.Controls.Add(this.urlBox);
            this.Controls.Add(this.WSDLURLLabel);
            this.Name = "MainUI";
            this.Text = "WSDL Describer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label WSDLURLLabel;
        private System.Windows.Forms.TextBox urlBox;
        private System.Windows.Forms.Label AuthorNameLabel;
        private System.Windows.Forms.TextBox authorNameBox;
        private System.Windows.Forms.Button generatorButton;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Timer timer;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.RichTextBox statusBox;
    }
}

