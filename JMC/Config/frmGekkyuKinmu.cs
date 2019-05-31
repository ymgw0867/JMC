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

namespace JMC.Config
{
    public partial class frmGekkyuKinmu : Form
    {
        public frmGekkyuKinmu(string pID)
        {
            InitializeComponent();

            adp.Fill(dts.社員ファイル);

            _grpID  = pID;      // グループID
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.社員ファイルTableAdapter adp = new DataSet1TableAdapters.社員ファイルTableAdapter();

        private void frmGekkyuKinmu_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);  // フォーム最大サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);  // フォーム最小サイズ

            foreach (var t in dts.社員ファイル.Where(a=>a.ID == Utility.StrtoInt(_grpID)))
            {
                _fName = t.ファイル名;     // ファイルパス名
                //_sheetName = t.シート名; // シート名
                _status = t.区分;     // 社員・パート区分
            }

            // パートタイマーコンボボックス
            if (_status == global.STATUS_PART)
            {
                Cursor = Cursors.WaitCursor;
                //Utility.ComboShain.xlsArrayLoad(_fName, _sheetName, comboBox1, global.flgOn);
                Utility.ComboShain.csvArrayLoad(_fName, comboBox1, global.flgOn);
                radioButton2.Checked = true;
                Cursor = Cursors.Default;
            }
            else
            {
                radioButton2.Checked = false;
            }

            GridViewSetting(dg1);                                   // グリッドビュー設定
            GridViewShow(dg1);                                      // グリッドビューへデータ表示
            this.Text = _grpName + " " + this.Text;                 // 画面キャプション
            DispClr();                                              // 画面初期化
        }

        // データ領域情報
        string _grpID = string.Empty;
        string _fName = string.Empty;
        string _sheetName = string.Empty;
        string _grpName = string.Empty;
        int _status = 0;

        // 登録モード
        //int _fMode = 0;

        // 選択データＩＤ
        string _SelectID = "0";

        // 登録内容
        string _Pt = string.Empty;
        string _BumonCode = string.Empty;   // 部門コード
        string _BumonName = string.Empty;   // 部門名

        // グリッドビューカラム名
        private string cCode = "c1";
        private string cName = "c2";
        private string cTime = "c3";
        private string cDate = "c4";
        private string cID = "c5";
        private string cbCode = "c6";
        private string cbName = "c7";

        Utility.xlsShain[] xS = null;


        ///------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューの定義を行います  </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------
        private void GridViewSetting(DataGridView tempDGV)
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
                //tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                tempDGV.Height = 312;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                //各列幅指定
                tempDGV.Columns.Add(cCode, "社員番号");
                tempDGV.Columns.Add(cName, "氏名");
                tempDGV.Columns.Add(cTime, "月間規定勤務時間");
                tempDGV.Columns.Add(cbCode, "コード");
                tempDGV.Columns.Add(cbName, "所属名");
                tempDGV.Columns.Add(cDate, "更新日");
                tempDGV.Columns.Add(cID, "ID");
                tempDGV.Columns[cID].Visible = false;

                tempDGV.Columns[cCode].Width = 100;
                tempDGV.Columns[cName].Width = 110;
                tempDGV.Columns[cTime].Width = 150;
                tempDGV.Columns[cbCode].Width = 80;
                tempDGV.Columns[cbName].Width = 160;
                tempDGV.Columns[cDate].Width = 110;
                tempDGV.Columns[cID].Width = 110;
                tempDGV.Columns[cbName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cbCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
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

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     データをグリッドビューへ表示します </summary>
        /// <param name="tempGrid">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------
        private void GridViewShow(DataGridView tempGrid)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            string mySql = "select * from 月給者勤務時間 ";
            mySql += "where データ領域 = '" + _grpID + "' "; 
            mySql += "order by 社員番号";

            sCom.Connection = Con.cnOpen();
            sCom.CommandText = mySql;
            dr = sCom.ExecuteReader();

            int iX = 0;
            int ms = 0;
            tempGrid.RowCount = 0;

            while (dr.Read())
            {
                tempGrid.Rows.Add();
                tempGrid[cCode, iX].Value = dr["社員番号"].ToString();
                tempGrid[cName, iX].Value = dr["氏名"].ToString();
                tempGrid[cTime, iX].Value = double.Parse(dr["月間勤務時間"].ToString()).ToString();
                tempGrid[cbCode, iX].Value = dr["所属コード"].ToString().ToString();
                tempGrid[cbName, iX].Value = dr["所属名"].ToString().ToString();
                tempGrid[cDate, iX].Value = DateTime.Parse(dr["更新年月日"].ToString()).ToShortDateString();
                tempGrid[cID, iX].Value = dr["ID"].ToString();

                // 月間勤務時間が0のとき赤表示
                if (double.Parse(dr["月間勤務時間"].ToString()) == 0)
                {
                    tempGrid.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;
                    ms++;
                }

                iX++;
            }

            dr.Close();
            sCom.Connection.Close();

            tempGrid.CurrentCell = null;

            // 未設定件数表示
            if (ms > 0) label5.Text = "月間規定勤務時間が未設定のデータが " + ms.ToString() + "件あります";
            else label5.Text = string.Empty; ;
        }

        private void DispClr()
        {
            lblNumber.Text = string.Empty;
            lblName.Text = string.Empty;
            textBox2.Text = string.Empty;

            button1.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;
            button5.Enabled = false;

            dg1.CurrentCell = null;

            //radioButton1.Checked = true;
            //radioButton2.Checked = false;
            _Pt = string.Empty;
            _BumonCode = string.Empty;
            _BumonName = string.Empty;
            groupBox1.Enabled = true;

            // パートタイマーコンボボックス
            if (_status == global.STATUS_PART)
            {
                radioButton2.Checked = true;
                radioButton2.Enabled = true;
                comboBox1.Enabled = true;
                radioButton1.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = true;
                button5.Enabled = true;
            }
            else
            {
                radioButton1.Checked = true;
                radioButton1.Enabled = true;
                button1.Enabled = true;
                radioButton2.Enabled = false;
                comboBox1.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("社員情報をインポートします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;

            // 社員情報インポート実行
            Cursor = Cursors.WaitCursor;
            MasterImportCsv(_fName);
            Cursor = Cursors.Default;
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     社員情報インポート実行 </summary>
        ///---------------------------------------------------------
        private void MasterImportCsv(string fName)
        {
            Utility.ComboShain.csvArrayLoad(fName, ref xS);

            int impCnt = 0;

            try
            {
                foreach (var t in xS.Where(a => a.sCode > 1 && a.kbn == 1).OrderBy(a => a.sCode))
                {
                    // データグリッドに表示されているか調べる
                    bool match = false;
                    for (int i = 0; i < dg1.Rows.Count; i++)
                    {
                        if (t.sCode == Utility.StrtoInt(dg1[cCode, i].Value.ToString()))
                        {
                            match = true;
                            break;
                        }
                    }

                    // データグリッドに表示されていなければインポート処理実行
                    if (!match)
                    {
                        recAdd("0", t.sCode.ToString(), t.sName, t.bCode.ToString(), t.bName);
                        impCnt++;
                    }
                }

                // データグリッド再表示
                GridViewShow(dg1);

                // 結果
                string resultMsg = string.Empty;

                if (impCnt > 0)
                {
                    resultMsg = impCnt.ToString() + "件の社員情報が登録されました。登録された社員は各々規定勤務時間を入力して下さい。";
                }
                else
                {
                    resultMsg = "新規に登録された社員情報はありませんでした";
                }

                MessageBox.Show(resultMsg, "インポート結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     社員情報インポート実行 </summary>
        ///---------------------------------------------------------
        private void MasterImportXls(string fName, string sheetName)
        {
            Utility.ComboShain.xlsArrayLoad(fName, sheetName, ref xS);

            int impCnt = 0;

            try
            {
                foreach (var t in xS.Where(a => a.sCode > 1).OrderBy(a => a.sCode))
                {
                    // データグリッドに表示されているか調べる
                    bool match = false;
                    for (int i = 0; i < dg1.Rows.Count; i++)
                    {
                        if (t.sCode == Utility.StrtoInt(dg1[cCode, i].Value.ToString()))
                        {
                            match = true;
                            break;
                        }
                    }

                    // データグリッドに表示されていなければインポート処理実行
                    if (!match)
                    {
                        recAdd("0", t.sCode.ToString(), t.sName, t.bCode.ToString(), t.bName);
                        impCnt++;
                    }
                }

                // データグリッド再表示
                GridViewShow(dg1);

                // 結果
                string resultMsg = string.Empty;

                if (impCnt > 0)
                {
                    resultMsg = impCnt.ToString() + "件の社員情報が登録されました。登録された社員は各々規定勤務時間を入力して下さい。";
                }
                else
                {
                    resultMsg = "新規に登録された社員情報はありませんでした";
                }

                MessageBox.Show(resultMsg, "インポート結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     月間勤務時間設定データを新規登録する </summary>
        /// <param name="sID">
        ///     ID</param>
        /// <param name="sCode">
        ///     社員コード</param>
        /// <param name="sSei">
        ///     社員姓</param>
        /// <param name="sMEi">
        ///     社員名</param>
        /// <param name="bCode">
        ///     所属コード</param>
        /// <param name="bName">
        ///     所属名</param>
        ///-----------------------------------------------------------
        private void recAdd(string sID, string sCode, string sSname, string bCode, string bName)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            try
            {
                sCom.Connection = Con.cnOpen();

                string mySql = "insert into 月給者勤務時間 (";
                mySql += "データ領域, ID, 社員番号, 氏名, 月間勤務時間, 所属コード, 所属名, 更新年月日) ";
                mySql += "values (?,?,?,?,?,?,?,?)";

                sCom.CommandText = mySql;
                sCom.Parameters.AddWithValue("@db", _grpID);
                sCom.Parameters.AddWithValue("@ID", sID);
                sCom.Parameters.AddWithValue("@Code", sCode);
                sCom.Parameters.AddWithValue("@Name", sSname);
                sCom.Parameters.AddWithValue("@Time", "0");
                sCom.Parameters.AddWithValue("@bCode", bCode);
                sCom.Parameters.AddWithValue("@bName", bName);
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
            button1.Enabled = false;
            comboBox1.Enabled = false;
            button2.Enabled = true;
            button3.Enabled = true;
            button5.Enabled = true;

            groupBox1.Enabled = false;
            radioButton1.Enabled = false;
            radioButton2.Enabled = false;
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     選択データ表示 </summary>
        /// <param name="r">
        ///     グリッドビュー行Index</param>
        ///-----------------------------------------------------------
        private void GetGridData(int r)
        {
            _SelectID = dg1[cCode, r].Value.ToString(); // 2016/11/21

            lblNumber.Text = dg1[cCode, r].Value.ToString();
            lblName.Text = dg1[cName, r].Value.ToString();
            textBox2.Text = dg1[cTime, r].Value.ToString();
            textBox2.Focus();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(lblName.Text + "の規定勤務時間を更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                return;

            // 2017/06/21 撤廃
            //if (Utility.StrtoInt(Utility.NulltoStr(textBox2.Text)) == 0)
            //{
            //    MessageBox.Show("勤務時間を入力して下さい","勤務時間未入力",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            //    return;
            //}

            // データ更新
            if (_Pt == string.Empty)
            {
                dataUpdate();
            }
            else
            {
                PartDataUpdate();
            }

            // データグリッド再表示
            GridViewShow(dg1);

            // 画面初期化
            DispClr();
        }

        /// <summary>
        /// データ更新
        /// </summary>
        private void dataUpdate()
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            try
            {
                sCom.Connection = Con.cnOpen();

                string mySql = "update 月給者勤務時間 set ";
                mySql += "月間勤務時間=?, 更新年月日=? ";
                mySql += "where 社員番号 = ?";

                sCom.CommandText = mySql;
                sCom.Parameters.AddWithValue("@Time", textBox2.Text);
                sCom.Parameters.AddWithValue("@Date", DateTime.Today);
                sCom.Parameters.AddWithValue("@ID", lblNumber.Text);

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

        /// <summary>
        /// パートデータ追加登録
        /// </summary>
        private void PartDataUpdate()
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            try
            {
                sCom.Connection = Con.cnOpen();

                string mySql = "insert into 月給者勤務時間 (";
                mySql += "データ領域, ID, 社員番号, 氏名, 月間勤務時間, 所属コード, 所属名, 更新年月日) ";
                mySql += "values (?,?,?,?,?,?,?,?)";

                sCom.CommandText = mySql;
                sCom.Parameters.AddWithValue("@db", _grpID);
                sCom.Parameters.AddWithValue("@ID", _SelectID);
                sCom.Parameters.AddWithValue("@Code", lblNumber.Text);
                sCom.Parameters.AddWithValue("@Name", lblName.Text);
                sCom.Parameters.AddWithValue("@Time", textBox2.Text);
                sCom.Parameters.AddWithValue("@bCode", _BumonCode);
                sCom.Parameters.AddWithValue("@bName", _BumonName);
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

        private void button5_Click(object sender, EventArgs e)
        {
            DispClr();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(lblName.Text + "を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;

            // データ削除
            dataDel();

            // 画面再表示
            GridViewShow(dg1);

            // 画面初期化
            DispClr();
        }

        private void dataDel()
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            try
            {
                sCom.Connection = Con.cnOpen();
                sCom.CommandText = "delete from 月給者勤務時間 where 社員番号 = ?";
                sCom.Parameters.AddWithValue("@ID", _SelectID);

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

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == string.Empty) textBox2.Text = "0";
        }

        private void dg1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1) return;
            Utility.ComboShain cmb = (Utility.ComboShain)comboBox1.SelectedItem;

            // データグリッドに表示されているか調べる
            int match = 0;
            for (int i = 0; i < dg1.Rows.Count; i++)
            {
                if (cmb.ID.ToString() == dg1[cID, i].Value.ToString() && cmb.code == dg1[cCode, i].Value.ToString())
                {
                    match = 1;
                    break;
                }
            }

            if (match == 1)
            {
                MessageBox.Show("既に登録済みです", "パートタイマー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            lblNumber.Text = cmb.code;
            lblName.Text = cmb.Name;
            _SelectID = cmb.ID.ToString();
            _BumonCode = cmb.BumonCode;
            _BumonName = cmb.BumonName;
            _Pt = cmb.code;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                button1.Enabled = true;
                radioButton2.Checked = false;
                comboBox1.Enabled = false;
            }
            else
            {
                button1.Enabled = false;
                radioButton2.Checked = true;
                comboBox1.Enabled = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                comboBox1.Enabled = true;
                radioButton1.Checked = false;
                button1.Enabled = false;
                button2.Enabled = true;
                button5.Enabled = true;
            }
            else
            {
                button1.Enabled = true;
                radioButton2.Checked = false;
                comboBox1.Enabled = false;
            }
        }
    }
}
