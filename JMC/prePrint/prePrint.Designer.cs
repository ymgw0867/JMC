namespace JMC.prePrint
{
    partial class prePrint
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(prePrint));
            this.label2 = new System.Windows.Forms.Label();
            this.dg1 = new System.Windows.Forms.DataGridView();
            this.btnPrn = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.txtSNo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnSel = new System.Windows.Forms.Button();
            this.txtENo = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnCheckOn = new System.Windows.Forms.Button();
            this.btnCheckOff = new System.Windows.Forms.Button();
            this.cmbBumonS = new System.Windows.Forms.ComboBox();
            this.cmbBumonE = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.SteelBlue;
            this.label2.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.Window;
            this.label2.Location = new System.Drawing.Point(254, 20);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 28);
            this.label2.TabIndex = 5;
            this.label2.Text = "所属範囲";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dg1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.dg1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg1.Location = new System.Drawing.Point(10, 126);
            this.dg1.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.dg1.Name = "dg1";
            this.dg1.ReadOnly = true;
            this.dg1.RowTemplate.Height = 21;
            this.dg1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dg1.Size = new System.Drawing.Size(1110, 521);
            this.dg1.TabIndex = 12;
            this.dg1.TabStop = false;
            // 
            // btnPrn
            // 
            this.btnPrn.Font = new System.Drawing.Font("メイリオ", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPrn.Location = new System.Drawing.Point(916, 661);
            this.btnPrn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnPrn.Name = "btnPrn";
            this.btnPrn.Size = new System.Drawing.Size(99, 42);
            this.btnPrn.TabIndex = 11;
            this.btnPrn.Text = "印刷(&P)";
            this.btnPrn.UseVisualStyleBackColor = true;
            this.btnPrn.Click += new System.EventHandler(this.btnPrn_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Font = new System.Drawing.Font("メイリオ", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(1021, 661);
            this.btnRtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(99, 42);
            this.btnRtn.TabIndex = 12;
            this.btnRtn.Text = "終了(&E)";
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // txtSNo
            // 
            this.txtSNo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtSNo.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtSNo.Location = new System.Drawing.Point(325, 52);
            this.txtSNo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtSNo.MaxLength = 5;
            this.txtSNo.Name = "txtSNo";
            this.txtSNo.Size = new System.Drawing.Size(60, 27);
            this.txtSNo.TabIndex = 5;
            this.txtSNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtSNo.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtSNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.SteelBlue;
            this.label1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.SystemColors.Window;
            this.label1.Location = new System.Drawing.Point(254, 52);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 27);
            this.label1.TabIndex = 10;
            this.label1.Text = "社員番号";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtYear
            // 
            this.txtYear.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtYear.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtYear.Location = new System.Drawing.Point(66, 20);
            this.txtYear.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtYear.MaxLength = 2;
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(42, 31);
            this.txtYear.TabIndex = 0;
            this.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtYear.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtYear.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.SteelBlue;
            this.label3.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.SystemColors.Window;
            this.label3.Location = new System.Drawing.Point(108, 20);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 31);
            this.label3.TabIndex = 12;
            this.label3.Text = "年";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtMonth
            // 
            this.txtMonth.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMonth.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtMonth.Location = new System.Drawing.Point(136, 20);
            this.txtMonth.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtMonth.MaxLength = 2;
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(42, 31);
            this.txtMonth.TabIndex = 1;
            this.txtMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMonth.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtMonth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.SteelBlue;
            this.label4.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.SystemColors.Window;
            this.label4.Location = new System.Drawing.Point(178, 20);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(28, 31);
            this.label4.TabIndex = 14;
            this.label4.Text = "月";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.SteelBlue;
            this.label5.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.SystemColors.Window;
            this.label5.Location = new System.Drawing.Point(15, 20);
            this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 31);
            this.label5.TabIndex = 15;
            this.label5.Text = "２０";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnSel
            // 
            this.btnSel.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnSel.Location = new System.Drawing.Point(1032, 19);
            this.btnSel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSel.Name = "btnSel";
            this.btnSel.Size = new System.Drawing.Size(84, 34);
            this.btnSel.TabIndex = 7;
            this.btnSel.Text = "検索(&S)";
            this.btnSel.UseVisualStyleBackColor = true;
            this.btnSel.Click += new System.EventHandler(this.btnSel_Click);
            // 
            // txtENo
            // 
            this.txtENo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtENo.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtENo.Location = new System.Drawing.Point(412, 52);
            this.txtENo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtENo.MaxLength = 5;
            this.txtENo.Name = "txtENo";
            this.txtENo.Size = new System.Drawing.Size(60, 27);
            this.txtENo.TabIndex = 6;
            this.txtENo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtENo.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtENo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Control;
            this.label6.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(624, 20);
            this.label6.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(23, 33);
            this.label6.TabIndex = 19;
            this.label6.Text = "～";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Control;
            this.label7.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(386, 52);
            this.label7.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(26, 33);
            this.label7.TabIndex = 20;
            this.label7.Text = "～";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnCheckOn
            // 
            this.btnCheckOn.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCheckOn.Location = new System.Drawing.Point(19, 661);
            this.btnCheckOn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnCheckOn.Name = "btnCheckOn";
            this.btnCheckOn.Size = new System.Drawing.Size(144, 42);
            this.btnCheckOn.TabIndex = 9;
            this.btnCheckOn.Text = "全てチェック(&N)";
            this.btnCheckOn.UseVisualStyleBackColor = true;
            this.btnCheckOn.Click += new System.EventHandler(this.btnCheckOn_Click);
            // 
            // btnCheckOff
            // 
            this.btnCheckOff.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCheckOff.Location = new System.Drawing.Point(169, 661);
            this.btnCheckOff.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnCheckOff.Name = "btnCheckOff";
            this.btnCheckOff.Size = new System.Drawing.Size(183, 42);
            this.btnCheckOff.TabIndex = 10;
            this.btnCheckOff.Text = "全てのチェックをはずす(&F)";
            this.btnCheckOff.UseVisualStyleBackColor = true;
            this.btnCheckOff.Click += new System.EventHandler(this.btnCheckOff_Click);
            // 
            // cmbBumonS
            // 
            this.cmbBumonS.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbBumonS.FormattingEnabled = true;
            this.cmbBumonS.Location = new System.Drawing.Point(325, 20);
            this.cmbBumonS.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cmbBumonS.Name = "cmbBumonS";
            this.cmbBumonS.Size = new System.Drawing.Size(301, 28);
            this.cmbBumonS.TabIndex = 3;
            // 
            // cmbBumonE
            // 
            this.cmbBumonE.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbBumonE.FormattingEnabled = true;
            this.cmbBumonE.Location = new System.Drawing.Point(644, 19);
            this.cmbBumonE.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cmbBumonE.Name = "cmbBumonE";
            this.cmbBumonE.Size = new System.Drawing.Size(301, 28);
            this.cmbBumonE.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Location = new System.Drawing.Point(18, 57);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(188, 54);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "発行モード";
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radioButton2.Location = new System.Drawing.Point(102, 21);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(58, 27);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "個別";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radioButton1.Location = new System.Drawing.Point(29, 21);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(58, 27);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "通常";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.Color.SteelBlue;
            this.label8.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.ForeColor = System.Drawing.SystemColors.Window;
            this.label8.Location = new System.Drawing.Point(254, 84);
            this.label8.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(71, 28);
            this.label8.TabIndex = 22;
            this.label8.Text = "社員選択";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(325, 84);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(239, 28);
            this.comboBox1.TabIndex = 8;
            // 
            // comboBox2
            // 
            this.comboBox2.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(644, 84);
            this.comboBox2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(301, 28);
            this.comboBox2.TabIndex = 9;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.Color.SteelBlue;
            this.label9.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.ForeColor = System.Drawing.SystemColors.Window;
            this.label9.Location = new System.Drawing.Point(573, 84);
            this.label9.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(71, 28);
            this.label9.TabIndex = 25;
            this.label9.Text = "所属";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnAdd
            // 
            this.btnAdd.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnAdd.Location = new System.Drawing.Point(1032, 80);
            this.btnAdd.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(84, 34);
            this.btnAdd.TabIndex = 10;
            this.btnAdd.Text = "追加(&D)";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.radioButton3);
            this.groupBox2.Controls.Add(this.radioButton4);
            this.groupBox2.Location = new System.Drawing.Point(376, 656);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(224, 47);
            this.groupBox2.TabIndex = 26;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "印刷順";
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radioButton3.Location = new System.Drawing.Point(127, 17);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(79, 24);
            this.radioButton3.TabIndex = 1;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "勤務先順";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radioButton4.Location = new System.Drawing.Point(18, 17);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(92, 24);
            this.radioButton4.TabIndex = 0;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "社員番号順";
            this.radioButton4.UseVisualStyleBackColor = true;
            // 
            // prePrint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1134, 714);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.cmbBumonE);
            this.Controls.Add(this.cmbBumonS);
            this.Controls.Add(this.btnCheckOff);
            this.Controls.Add(this.btnCheckOn);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtENo);
            this.Controls.Add(this.btnSel);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtMonth);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtSNo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnPrn);
            this.Controls.Add(this.dg1);
            this.Controls.Add(this.label2);
            this.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.Name = "prePrint";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Form1_KeyDown);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Form1_KeyPress);
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dg1;
        private System.Windows.Forms.Button btnPrn;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.TextBox txtSNo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnSel;
        private System.Windows.Forms.TextBox txtENo;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnCheckOn;
        private System.Windows.Forms.Button btnCheckOff;
        private System.Windows.Forms.ComboBox cmbBumonS;
        private System.Windows.Forms.ComboBox cmbBumonE;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.RadioButton radioButton4;
    }
}

