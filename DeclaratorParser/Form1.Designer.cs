namespace DeclaratorParser
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.btnLoadWord = new System.Windows.Forms.Button();
			this.btnTest = new System.Windows.Forms.Button();
			this.sStrip = new System.Windows.Forms.StatusStrip();
			this.tstripInfo = new System.Windows.Forms.ToolStripStatusLabel();
			this.btnWriteByRegion = new System.Windows.Forms.Button();
			this.btnTextStrangeOrgs = new System.Windows.Forms.Button();
			this.btnSetRegions = new System.Windows.Forms.Button();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.sStrip.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.SuspendLayout();
			// 
			// btnLoadWord
			// 
			this.btnLoadWord.Location = new System.Drawing.Point(33, 22);
			this.btnLoadWord.Name = "btnLoadWord";
			this.btnLoadWord.Size = new System.Drawing.Size(248, 23);
			this.btnLoadWord.TabIndex = 2;
			this.btnLoadWord.Text = "1. Загрузить декларацию (doc)";
			this.btnLoadWord.UseVisualStyleBackColor = true;
			this.btnLoadWord.Click += new System.EventHandler(this.btnLoadWord_Click);
			// 
			// btnTest
			// 
			this.btnTest.Location = new System.Drawing.Point(329, 190);
			this.btnTest.Name = "btnTest";
			this.btnTest.Size = new System.Drawing.Size(248, 23);
			this.btnTest.TabIndex = 3;
			this.btnTest.Text = "Тест парсера недвижимости";
			this.btnTest.UseVisualStyleBackColor = true;
			this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
			// 
			// sStrip
			// 
			this.sStrip.ImageScalingSize = new System.Drawing.Size(18, 18);
			this.sStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tstripInfo});
			this.sStrip.Location = new System.Drawing.Point(0, 283);
			this.sStrip.Name = "sStrip";
			this.sStrip.Size = new System.Drawing.Size(619, 22);
			this.sStrip.TabIndex = 4;
			// 
			// tstripInfo
			// 
			this.tstripInfo.Name = "tstripInfo";
			this.tstripInfo.Size = new System.Drawing.Size(0, 17);
			// 
			// btnWriteByRegion
			// 
			this.btnWriteByRegion.Location = new System.Drawing.Point(30, 190);
			this.btnWriteByRegion.Name = "btnWriteByRegion";
			this.btnWriteByRegion.Size = new System.Drawing.Size(248, 23);
			this.btnWriteByRegion.TabIndex = 6;
			this.btnWriteByRegion.Text = "3. Записать вузы по регионам";
			this.btnWriteByRegion.UseVisualStyleBackColor = true;
			this.btnWriteByRegion.Click += new System.EventHandler(this.btnWriteByRegion_Click);
			// 
			// btnTextStrangeOrgs
			// 
			this.btnTextStrangeOrgs.Location = new System.Drawing.Point(329, 99);
			this.btnTextStrangeOrgs.Name = "btnTextStrangeOrgs";
			this.btnTextStrangeOrgs.Size = new System.Drawing.Size(248, 23);
			this.btnTextStrangeOrgs.TabIndex = 7;
			this.btnTextStrangeOrgs.Text = "Тест неопознанных организаций";
			this.btnTextStrangeOrgs.UseVisualStyleBackColor = true;
			this.btnTextStrangeOrgs.Click += new System.EventHandler(this.btnTextStrangeOrgs_Click);
			// 
			// btnSetRegions
			// 
			this.btnSetRegions.Location = new System.Drawing.Point(33, 99);
			this.btnSetRegions.Name = "btnSetRegions";
			this.btnSetRegions.Size = new System.Drawing.Size(248, 26);
			this.btnSetRegions.TabIndex = 8;
			this.btnSetRegions.Text = "2. Прописать регионы";
			this.btnSetRegions.UseVisualStyleBackColor = true;
			this.btnSetRegions.Click += new System.EventHandler(this.btnSetRegions_Click);
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.InitialImage")));
			this.pictureBox1.Location = new System.Drawing.Point(329, 22);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(248, 34);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 9;
			this.pictureBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(33, 52);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(248, 44);
			this.label1.TabIndex = 10;
			this.label1.Text = "Парсер информации из декларации в формат Заполнятора (xml)";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(30, 129);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(248, 58);
			this.label2.TabIndex = 10;
			this.label2.Text = "Ищет вузы в реестре лицензий Рособрнадзора и вписывает в теги Organization и Regi" +
    "on в xml Заполнятора";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(27, 216);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(248, 52);
			this.label3.TabIndex = 10;
			this.label3.Text = "Создает отдельные xml (формат Заполнятора) для регионов и раскладывает по ним вуз" +
    "ы";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(329, 128);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(248, 52);
			this.label4.TabIndex = 10;
			this.label4.Text = "Отладчик для изучения проблемных вузов, которые не найдены в реестре Рособрнадзор" +
    "а";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(329, 216);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(248, 52);
			this.label5.TabIndex = 10;
			this.label5.Text = "Старый тест парсера недвижимости для Заполнятора. Не используется";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(619, 305);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.btnSetRegions);
			this.Controls.Add(this.btnTextStrangeOrgs);
			this.Controls.Add(this.btnWriteByRegion);
			this.Controls.Add(this.sStrip);
			this.Controls.Add(this.btnTest);
			this.Controls.Add(this.btnLoadWord);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "Form1";
			this.Text = "Парсер деклараций 1.0";
			this.sStrip.ResumeLayout(false);
			this.sStrip.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

		private System.Windows.Forms.Button btnLoadWord;
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.StatusStrip sStrip;
		private System.Windows.Forms.ToolStripStatusLabel tstripInfo;
		private System.Windows.Forms.Button btnWriteByRegion;
		private System.Windows.Forms.Button btnTextStrangeOrgs;
		private System.Windows.Forms.Button btnSetRegions;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
    }
}

