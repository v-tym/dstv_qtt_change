
namespace dstv_qtt_change
{
    partial class Form1
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.openExel = new System.Windows.Forms.Button();
            this.pathToDSTV = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.Start = new System.Windows.Forms.Button();
            this.messages = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox1.Location = new System.Drawing.Point(13, 13);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(563, 30);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "Путь к ексель файлу";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox2.Location = new System.Drawing.Point(13, 77);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(563, 30);
            this.textBox2.TabIndex = 1;
            this.textBox2.Text = "Путь к папке дств";
            // 
            // openExel
            // 
            this.openExel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.openExel.Location = new System.Drawing.Point(582, 6);
            this.openExel.Name = "openExel";
            this.openExel.Size = new System.Drawing.Size(201, 40);
            this.openExel.TabIndex = 2;
            this.openExel.Text = "Open excel";
            this.openExel.UseVisualStyleBackColor = true;
            this.openExel.Click += new System.EventHandler(this.button1_Click);
            // 
            // pathToDSTV
            // 
            this.pathToDSTV.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pathToDSTV.Location = new System.Drawing.Point(582, 71);
            this.pathToDSTV.Name = "pathToDSTV";
            this.pathToDSTV.Size = new System.Drawing.Size(201, 38);
            this.pathToDSTV.TabIndex = 4;
            this.pathToDSTV.Text = "Open path to DSTV";
            this.pathToDSTV.UseVisualStyleBackColor = true;
            this.pathToDSTV.Click += new System.EventHandler(this.pathToDSTV_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Start
            // 
            this.Start.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Start.Location = new System.Drawing.Point(13, 143);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(770, 45);
            this.Start.TabIndex = 5;
            this.Start.Text = "Start";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // messages
            // 
            this.messages.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.messages.Location = new System.Drawing.Point(13, 216);
            this.messages.Multiline = true;
            this.messages.Name = "messages";
            this.messages.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.messages.Size = new System.Drawing.Size(770, 174);
            this.messages.TabIndex = 6;
            this.messages.Text = "Сообщения ошибок...";
            this.messages.TextChanged += new System.EventHandler(this.messages_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(795, 402);
            this.Controls.Add(this.messages);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.pathToDSTV);
            this.Controls.Add(this.openExel);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button openExel;
        private System.Windows.Forms.Button pathToDSTV;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.TextBox messages;
    }
}

