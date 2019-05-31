namespace JMC.prePrint
{
    partial class frmPreCompany
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPreCompany));
            this.dg1 = new System.Windows.Forms.DataGridView();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).BeginInit();
            this.SuspendLayout();
            // 
            // dg1
            // 
            this.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg1.Location = new System.Drawing.Point(12, 12);
            this.dg1.Name = "dg1";
            this.dg1.RowTemplate.Height = 21;
            this.dg1.Size = new System.Drawing.Size(477, 305);
            this.dg1.TabIndex = 0;
            this.dg1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellClick);
            this.dg1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellDoubleClick);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(237, 381);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(80, 33);
            this.button2.TabIndex = 8;
            this.button2.Text = "登録(&U)";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(323, 381);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(80, 33);
            this.button3.TabIndex = 9;
            this.button3.Text = "削除(&D)";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(409, 381);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(80, 33);
            this.button4.TabIndex = 10;
            this.button4.Text = "終了(&E)";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(140, 327);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(349, 28);
            this.comboBox1.TabIndex = 15;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 330);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 20);
            this.label1.TabIndex = 16;
            this.label1.Text = "新規登録領域選択：";
            // 
            // frmPreCompany
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 426);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.dg1);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmPreCompany";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "出勤簿の有休記号等非表示にするデータ領域登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGekkyuKinmu_FormClosing);
            this.Load += new System.EventHandler(this.frmGekkyuKinmu_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dg1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
    }
}