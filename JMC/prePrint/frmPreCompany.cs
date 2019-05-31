using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JMC.Common;
using System.Data.OleDb;

namespace JMC.prePrint
{
    public partial class frmPreCompany : Form
    {
        public frmPreCompany()
        {
            InitializeComponent();
        }

        private void frmGekkyuKinmu_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);  // フォーム最大サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);  // フォーム最小サイズ
            GridViewSetting(dg1);                                   // グリッドビュー設定
            GridViewShow(dg1);                                      // グリッドビューへデータ表示
            Utility.ComboDataArea.load(comboBox1);
            DispClr();                                              // 画面初期化
        }

        // データ領域情報
        string _dbName = string.Empty;
        string _comName = string.Empty;

        // 登録モード
        int _fMode = 0;

        // 選択データＩＤ
        string _SelectID = "0";

        // グリッドビューカラム名
        private string cCode = "c1";
        private string cName = "c2";
        private string cTime = "c3";
        private string cDate = "c4";
        private string cID = "c5";
        private string cbCode = "c6";
        private string cbName = "c7";

        /// <summary>
        /// グリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 10, FontStyle.Regular);

                // 行の高さ
                //tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 305;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                //各列幅指定
                tempDGV.Columns.Add(cCode, "コード");
                tempDGV.Columns.Add(cName, "データ領域名");
                tempDGV.Columns.Add(cDate, "登録日");
                tempDGV.Columns.Add(cID, "ID");
                tempDGV.Columns[cID].Visible = false;

                tempDGV.Columns[cCode].Width = 100;
                tempDGV.Columns[cName].Width = 300;
                tempDGV.Columns[cDate].Width = 110;
                tempDGV.Columns[cID].Width = 110;
                tempDGV.Columns[cName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cID].Visible = false;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = false;

                //// ソート禁止
                //foreach (DataGridViewColumn c in tempDGV.Columns)
                //{
                //    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
                //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// データをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewShow(DataGridView tempGrid)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            string mySql = "select * from 有休非表示領域 ";

            sCom.Connection = Con.cnOpen();
            sCom.CommandText = mySql;
            dr = sCom.ExecuteReader();

            int iX = 0;
            int ms = 0;
            tempGrid.RowCount = 0;

            while (dr.Read())
            {
                tempGrid.Rows.Add();
                tempGrid[cCode, iX].Value = dr["CompanyCode"].ToString();
                tempGrid[cName, iX].Value = dr["CompanyName"].ToString();
                tempGrid[cDate, iX].Value = DateTime.Parse(dr["更新年月日"].ToString()).ToShortDateString();
                tempGrid[cID, iX].Value = dr["Name"].ToString();
                iX++;
            }

            dr.Close();
            sCom.Connection.Close();

            tempGrid.CurrentCell = null;
        }

        private void DispClr()
        {
            button2.Enabled = false;
            button3.Enabled = false;

            dg1.CurrentCell = null;
        }

        private void recAdd(string Name, string CompanyCode, string CompanyName)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            try
            {
                sCom.Connection = Con.cnOpen();

                string mySql = "insert into 有休非表示領域 (";
                mySql += "Name, CompanyCode, CompanyName, 更新年月日) ";
                mySql += "values (?,?,?,?)";

                sCom.CommandText = mySql;
                sCom.Parameters.AddWithValue("@Name", Name);
                sCom.Parameters.AddWithValue("@Code", CompanyCode);
                sCom.Parameters.AddWithValue("@comName", CompanyName);
                sCom.Parameters.AddWithValue("@Date", DateTime.Today);

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close(); 
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmGekkyuKinmu_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dg1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // 選択データ表示
            GetGridData(e.RowIndex);

            // ボタン
            button2.Enabled = true;
            button3.Enabled = true;
        }

        /// <summary>
        /// 選択データ表示
        /// </summary>
        /// <param name="r">グリッドビュー行Index</param>
        private void GetGridData(int r)
        {
            _SelectID = dg1[cID, r].Value.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == string.Empty || comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("新規に登録するデータ領域を選択して下さい", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
 
            if (MessageBox.Show(comboBox1.Text + "を登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                return;

            Utility.ComboDataArea cc = (Utility.ComboDataArea)comboBox1.SelectedItem;

            // データ更新
            recAdd(cc.ID, cc.code, cc.DisplayName);

            // データグリッド再表示
            GridViewShow(dg1);

            // 画面初期化
            DispClr();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DispClr();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dg1.SelectedRows.Count == 0)
            {
                MessageBox.Show("削除するデータ領域を選択して下さい", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
 
            if (MessageBox.Show(dg1[1, dg1.SelectedRows[0].Index].Value.ToString() + "を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;

            // データ削除
            dataDel(dg1[3, dg1.SelectedRows[0].Index].Value.ToString());

            // 画面再表示
            GridViewShow(dg1);

            // 画面初期化
            DispClr();
        }

        private void dataDel(string dName)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            try
            {
                sCom.Connection = Con.cnOpen();
                sCom.CommandText = "delete from 有休非表示領域 where Name = ?";
                sCom.Parameters.AddWithValue("@ID", dName);

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1) button2.Enabled = true;
            else button2.Enabled = false;
        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            button3.Enabled = true;
        }
    }
}
