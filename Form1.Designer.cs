namespace Excel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnVTOkuYaz = new Button();
            richTextBox1 = new RichTextBox();
            richTextBox2 = new RichTextBox();
            btnExcelOkuYaz = new Button();
            SuspendLayout();
            // 
            // btnVTOkuYaz
            // 
            btnVTOkuYaz.Location = new Point(25, 66);
            btnVTOkuYaz.Name = "btnVTOkuYaz";
            btnVTOkuYaz.Size = new Size(154, 63);
            btnVTOkuYaz.TabIndex = 0;
            btnVTOkuYaz.Text = "Veri Tabanı Oku ve Excel'e Yazdır";
            btnVTOkuYaz.UseVisualStyleBackColor = true;
            btnVTOkuYaz.Click += btnVTOkuma_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(244, 12);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(488, 117);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(244, 192);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(488, 120);
            richTextBox2.TabIndex = 2;
            richTextBox2.Text = "";
            // 
            // btnExcelOkuYaz
            // 
            btnExcelOkuYaz.Location = new Point(25, 247);
            btnExcelOkuYaz.Name = "btnExcelOkuYaz";
            btnExcelOkuYaz.Size = new Size(154, 65);
            btnExcelOkuYaz.TabIndex = 3;
            btnExcelOkuYaz.Text = "Excel'den Oku Veri Tabanına Yazdır";
            btnExcelOkuYaz.UseVisualStyleBackColor = true;
            btnExcelOkuYaz.Click += btnExcelOkuYaz_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(btnExcelOkuYaz);
            Controls.Add(richTextBox2);
            Controls.Add(richTextBox1);
            Controls.Add(btnVTOkuYaz);
            Name = "Form1";
            Text = "Veri Tabanı Excel Entegrasyonu";
            ResumeLayout(false);
        }

        #endregion

        private Button btnVTOkuYaz;
        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private Button btnExcelOkuYaz;
    }
}
