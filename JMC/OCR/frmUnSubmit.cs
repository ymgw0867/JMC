using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using JMC.Common;


namespace JMC.OCR
{
    public partial class frmUnSubmit : Form
    {
        public frmUnSubmit()
        {
            InitializeComponent();

            adp.Fill(dts.社員ファイル);
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.社員ファイルTableAdapter adp = new DataSet1TableAdapters.社員ファイルTableAdapter();
        
        private void frmUnSubmit_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // データ領域名コンボボックスのデータソースをセットする
            dataComboSet();

            // データグリッド定義
            GridviewSet(dataGridView1);

            // 画面初期化
            DispClear();
        }

        private void dataComboSet()
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            // データベースと接続する
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            OleDbDataReader dR = null;

            // 過去出勤簿ヘッダデータリーダーよりデータ領域名を取得する
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select distinct データ領域名 from 過去出勤簿ヘッダ ");
            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                // 給与グループ名を取得する  2016/11/23
                foreach (var t in dts.社員ファイル.Where(a => a.ID == Utility.StrtoInt(dR["データ領域名"].ToString())))
                {
                    if (t.Isグループ名Null())
                    {
                        comboBox1.Items.Add("");
                    }
                    else
                    {
                        comboBox1.Items.Add(t.グループ名);
                    }
                }
            }

            dR.Close();

            // 過去出勤簿ヘッダデータリーダーより所属名を取得する
            sb.Clear();
            sb.Append("select distinct 所属名 from 過去出勤簿ヘッダ ");
            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                comboBox2.Items.Add(dR["所属名"].ToString());
            }
            dR.Close();

            sCom.Connection.Close();
        }


        // カラム定義
        private string ColData = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColCode = "c3";
        private string ColName = "c4";
        //private string ColStatus = "c5";
        private string ColYear = "c6";
        private string ColMonth = "c7";
        private string ColID = "c8";

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        private void GridviewSet(DataGridView tempDGV)
        {
            try
            {
                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 10, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                tempDGV.Height = 643;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                // 各列幅指定
                tempDGV.Columns.Add(ColYear, "年");
                tempDGV.Columns.Add(ColMonth, "月");
                tempDGV.Columns.Add(ColData, "グループ");
                tempDGV.Columns.Add(ColSz, "所属コード");
                tempDGV.Columns.Add(ColSznm, "所属名");
                tempDGV.Columns.Add(ColCode, "個人番号");
                tempDGV.Columns.Add(ColName, "氏名");
                tempDGV.Columns.Add(ColID, "ID");

                tempDGV.Columns[ColYear].Width = 60;
                tempDGV.Columns[ColMonth].Width = 40;
                tempDGV.Columns[ColData].Width = 200;
                tempDGV.Columns[ColSz].Width = 90;
                tempDGV.Columns[ColSznm].Width = 200;
                tempDGV.Columns[ColCode].Width = 80;
                tempDGV.Columns[ColName].Width = 200;

                tempDGV.Columns[ColID].Visible = false;

                tempDGV.Columns[ColName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[ColYear].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColMonth].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColSz].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                tempDGV.ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 画面初期化
        /// </summary>
        private void DispClear()
        {
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            comboBox1.Text = string.Empty;
            comboBox1.SelectedIndex = -1;
            comboBox2.Text = string.Empty;
            comboBox2.SelectedIndex = -1;
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
            if (errCheck())
            {
                DataSelect();
            }
        }

        private bool errCheck()
        {
            try 
	        {
                if (txtYear.Text != string.Empty && !Utility.NumericCheck(txtYear.Text))
                {
                    txtYear.Focus();
                    throw new Exception("年が正しくありません");
                }

                if (txtMonth.Text != string.Empty && !Utility.NumericCheck(txtMonth.Text))
                {
                    txtMonth.Focus();
                    throw new Exception("月が正しくありません");
                }

                if (txtMonth.Text != string.Empty)
                {
                    if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
                    {
                        txtMonth.Focus();
                        throw new Exception("月が正しくありません");
                    }
                }
            }
	        catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
	        }
            return true;
        }

        private void DataSelect()
        {
            // データグリッドビューの表示を初期化する
            dataGridView1.RowCount = 0;

            // 給与グループテーブルからグループ名を取得する 2016/11/23
            string grpID = "";
            if (dts.社員ファイル.Any(a=>a.グループ名 == comboBox1.Text))
            {
                var s = dts.社員ファイル.Single(a => a.グループ名 == comboBox1.Text);
                grpID = s.ID.ToString();
            }
            
            // データベースと接続する
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            OleDbDataReader dR = null;

            // 過去出勤簿ヘッダデータリーダーを取得する
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select * from 過去出勤簿ヘッダ where ID <> '' ");

            // データ領域名
            if (comboBox1.SelectedIndex != -1) sb.Append("and データ領域名 = '" + grpID + "' ");
                
            // 年
            if (txtYear.Text != string.Empty) sb.Append("and 年 = " + txtYear.Text + " ");
                
            // 月
            if (txtMonth.Text != string.Empty) sb.Append("and 月 = " + txtMonth.Text + " ");

            // データ領域名
            if (comboBox2.SelectedIndex != -1) sb.Append("and 所属名 = '" + comboBox2.Text + "' ");
                
            sb.Append("order by 年, 月, データ領域名, 所属コード, 個人番号");

            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                // グリッドへ表示する
                gridShow(dR, dataGridView1);
            }
            dR.Close();
            sCom.Connection.Close();
            dataGridView1.CurrentCell = null;

            // 終了
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("該当するデータはありませんでした", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     データグリッドへ表示する </summary>
        /// <param name="sdR">
        ///     データリーダーオブジェクト</param>
        /// <param name="stutas">
        ///     ステータス</param>
        /// <param name="g">
        ///     datagridviewオブジェクト</param>
        ///-------------------------------------------------------------------
        private void gridShow(OleDbDataReader sdR, DataGridView g)
        {
            g.Rows.Add();

            // 給与グループテーブルからグループ名を取得する 2016/11/23
            string grpName = "";
            if (dts.社員ファイル.Any(a => a.ID == Utility.StrtoInt(sdR["データ領域名"].ToString())))
            {
                var s = dts.社員ファイル.Single(a => a.ID == Utility.StrtoInt(sdR["データ領域名"].ToString()));
                grpName = s.グループ名;
            }

            g[ColData, g.Rows.Count - 1].Value = grpName;
            g[ColYear, g.Rows.Count - 1].Value = sdR["年"].ToString();
            g[ColMonth, g.Rows.Count - 1].Value = sdR["月"].ToString();
            g[ColSz, g.Rows.Count - 1].Value = sdR["所属コード"].ToString().PadLeft(5, '0');
            g[ColSznm, g.Rows.Count - 1].Value = sdR["所属名"].ToString();
            g[ColCode, g.Rows.Count - 1].Value = sdR["個人番号"].ToString().PadLeft(5, '0');
            g[ColName, g.Rows.Count - 1].Value = sdR["氏名"].ToString();
            g[ColID, g.Rows.Count - 1].Value = sdR["ID"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmUnSubmit_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string rID = string.Empty;

            rID = dataGridView1[ColID, dataGridView1.SelectedRows[0].Index].Value.ToString();

            if (rID != string.Empty)
            {
                this.Hide();
                OCR.frmCorrect frm = new OCR.frmCorrect(string.Empty, string.Empty, 0, string.Empty, rID);
                frm.ShowDialog();
                this.Show();
            }
        }
    }
}
