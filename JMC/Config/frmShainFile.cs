using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JMC.Common;

namespace JMC.Config
{
    public partial class frmShainFile : Form
    {
        // マスター名
        string msName = "給与グループ";

        // フォームモードインスタンス
        Utility.frmMode fMode = new Utility.frmMode();

        // 居宅支援事業所マスターテーブルアダプター生成
        DataSet1TableAdapters.社員ファイルTableAdapter adp = new DataSet1TableAdapters.社員ファイルTableAdapter();

        // データテーブル生成
        DataSet1 dts = new DataSet1();

        public frmShainFile()
        {
            InitializeComponent();

            // データテーブルにデータを読み込む
            adp.Fill(dts.社員ファイル);
        }

        private void frm_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dg);

            // データグリッドビューにデータを表示します
            GridViewShow(dg);

            // 画面初期化
            DispInitial();
        }

        //カラム定義
        //string cCode = "col1";
        string cName = "col2";
        string cStatus = "col3";
        string cSheetNum = "col4";
        string cID = "col5";
        string cDate = "col6";
        string cGroup = "col7";


        ///-------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-------------------------------------------------------------------
        private void GridViewSetting(DataGridView g)
        {
            try
            {
                g.EnableHeadersVisualStyles = false;
                g.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                g.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                g.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                g.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                g.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                g.ColumnHeadersHeight = 20;
                g.RowTemplate.Height = 20;

                // 全体の高さ
                g.Height = 301;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
                
                g.Columns.Add(cGroup, "グループ名");
                g.Columns.Add(cStatus, "種別");
                g.Columns.Add(cName, "ファイル名");
                //g.Columns.Add(cSheetNum, "シート名");
                g.Columns.Add(cDate, "更新年月日");
                g.Columns.Add(cID, "cID");

                g.Columns[cGroup].Width = 160;
                g.Columns[cStatus].Width = 70;
                g.Columns[cName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                //g.Columns[cSheetNum].Width = 110;
                g.Columns[cDate].Width = 140;
                g.Columns[cID].Visible = false;

                //g.Columns[cSheetNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 行ヘッダを表示しない
                g.RowHeadersVisible = false;

                // 選択モード
                g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                g.MultiSelect = false;

                // 編集不可とする
                g.ReadOnly = true;

                // 追加行表示しない
                g.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                g.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                g.AllowUserToOrderColumns = false;

                // 列サイズ変更可
                g.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                g.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                g.StandardTab = true;

                // 罫線
                g.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                g.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューにデータを表示します </summary>
        /// <param name="tempGrid">
        ///     データグリッドビューオブジェクト名</param>
        ///---------------------------------------------------------------------
        private void GridViewShow(DataGridView g)
        {
            g.Rows.Clear();

            int iX = 0;
            global gl = new global();

            try
            {
                foreach (var t in dts.社員ファイル.OrderBy(a => a.ID))
                {
                    g.Rows.Add();

                    if (t.区分 == global.STATUS_SHAIN)
                    {
                        g[cStatus, iX].Value = "社員";
                    }
                    else if (t.区分 == global.STATUS_PART)
                    {
                        g[cStatus, iX].Value = "パート";
                    }

                    if (t.Isグループ名Null())
                    {
                        g[cGroup, iX].Value = "";
                    }
                    else
                    {
                        g[cGroup, iX].Value = t.グループ名;
                    }

                    g[cName, iX].Value = t.ファイル名;
                    //g[cSheetNum, iX].Value = t.シート名.ToString();
                    g[cDate, iX].Value = t.更新年月日;
                    g[cID, iX].Value = t.ID.ToString();

                    iX++;
                }

                if (g.Rows.Count > 0)
                {
                    g.CurrentCell = null;
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message); 
            }
        }

        ///--------------------------------------------------
        /// <summary>
        ///     画面の初期化 </summary>
        ///--------------------------------------------------
        private void DispInitial()
        {
            fMode.Mode = global.FORM_ADDMODE;
            txtGrpName.Text = string.Empty;
            comboBox1.SelectedIndex = -1;
            txtFileName.Text = string.Empty;
            //txtSheetNum.Text = string.Empty;

            linkLabel4.Enabled = true;
            linkLabel2.Enabled = false;
            linkLabel3.Enabled = false;
            txtGrpName.Focus();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
        }

        //登録データチェック
        private Boolean fDataCheck()
        {
            try
            {
                // グループ名チェック
                if (txtGrpName.Text == string.Empty)
                {
                    txtGrpName.Focus();
                    throw new Exception("グループ名を入力してください");
                }

                // 雇用種別
                if (comboBox1.SelectedIndex == -1 || comboBox1.Text == string.Empty)
                {
                    comboBox1.Focus();
                    throw new Exception("雇用種別を選択してください");
                }

                // 名称チェック
                if (txtFileName.Text.Trim() == string.Empty)
                {
                    txtFileName.Focus();
                    throw new Exception("ファイルを選択してください");
                }

                //// シート番号チェック
                //if (txtSheetNum.Text == string.Empty)
                //{
                //    txtFileName.Focus();
                //    throw new Exception("シート番号を入力してください");
                //}

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, msName + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     グリッドビュー行選択時処理　</summary>
        ///----------------------------------------------------------
        private void GridEnter()
        {
            string msgStr;
            fMode.rowIndex = dg.SelectedRows[0].Index;

            // 選択確認
            msgStr = "";
            msgStr += dg[0, fMode.rowIndex].Value.ToString() + "：" + dg[1, fMode.rowIndex].Value.ToString() + Environment.NewLine + Environment.NewLine;
            msgStr += "上記の" + msName + "が選択されました。よろしいですか？";

            if (MessageBox.Show(msgStr, "選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No) 
                return;

            // 対象となるデータテーブルROWを取得します
            DataSet1.社員ファイルRow sQuery = dts.社員ファイル.FindByID(int.Parse(dg[cID, fMode.rowIndex].Value.ToString()));

            if (sQuery != null)
            {
                // 編集画面に表示
                ShowData(sQuery);

                // モードステータスを「編集モード」にします
                fMode.Mode = global.FORM_EDITMODE;
            }
            else
            {
                MessageBox.Show(dg[0, fMode.rowIndex].Value.ToString() + "がキー不在です：データの読み込みに失敗しました", "データ取得エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// -------------------------------------------------------
        /// <summary>
        ///     マスターの内容を画面に表示する </summary>
        /// <param name="sTemp">
        ///     マスターインスタンス</param>
        /// -------------------------------------------------------
        private void ShowData(DataSet1.社員ファイルRow s)
        {
            fMode.ID = s.ID;

            if (s.Isグループ名Null())
            {
                txtGrpName.Text = string.Empty;
            }
            else
            {
                txtGrpName.Text = s.グループ名;
            }

            comboBox1.SelectedIndex = s.区分 - 1;
            //txtSheetNum.Text = s.シート名.ToString();
            txtFileName.Text = s.ファイル名;

            linkLabel2.Enabled = true;
            linkLabel3.Enabled = true;
        }

        private void dg_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // データセットの内容をデータベースへ反映させます
            adp.Update(dts.社員ファイル);

            this.Dispose();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
        }

        private void frmKintaiKbn_Shown(object sender, EventArgs e)
        {
            txtGrpName.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void txtCode_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void txtSh_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
        
        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //エラーチェック
            if (!fDataCheck()) return;

            switch (fMode.Mode)
            {
                // 新規登録
                case global.FORM_ADDMODE:

                    // 確認
                    if (MessageBox.Show("登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセットにデータを追加します
                    var s = dts.社員ファイル.New社員ファイルRow();
                    s.グループ名 = txtGrpName.Text;
                    s.ファイル名 = txtFileName.Text;
                    //s.シート名 = txtSheetNum.Text;
                    s.区分 = comboBox1.SelectedIndex + 1;
                    s.更新年月日 = DateTime.Now;

                    dts.社員ファイル.Add社員ファイルRow(s);

                    break;

                // 更新処理
                case global.FORM_EDITMODE:

                    // 確認
                    if (MessageBox.Show("更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセット更新
                    var r = dts.社員ファイル.Single(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                               a.ID == fMode.ID);

                    if (!r.HasErrors)
                    {
                        r.グループ名 = txtGrpName.Text;
                        r.ファイル名 = txtFileName.Text;
                        //r.シート名 = txtSheetNum.Text;
                        r.区分 = comboBox1.SelectedIndex + 1;
                        r.更新年月日 = DateTime.Now;
                    }
                    else
                    {
                        MessageBox.Show(fMode.ID + "がキー不在です：データの更新に失敗しました", "更新エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                    break;

                default:
                    break;
            }

            // 更新をコミット
            adp.Update(dts.社員ファイル);

            // データテーブルにデータを読み込む
            adp.Fill(dts.社員ファイル);

            // 画面データ消去
            DispInitial();

            // グリッド表示
            GridViewShow(dg);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                // 確認
                if (MessageBox.Show("削除してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                    return;

                // 削除データ取得（エラー回避のためDataRowState.Deleted と DataRowState.Detachedは除外して抽出する）
                var d = dts.社員ファイル.Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached && a.ID == fMode.ID);

                // foreach用の配列を作成する
                var list = d.ToList();

                // 削除
                foreach (var it in list)
                {
                    DataSet1.社員ファイルRow dl = dts.社員ファイル.FindByID(it.ID);
                    dl.Delete();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("データの削除に失敗しました" + Environment.NewLine + ex.Message);
            }
            finally
            {
                // 削除をコミット
                adp.Update(dts.社員ファイル);

                // データテーブルにデータを読み込む
                adp.Fill(dts.社員ファイル);

                // 画面データ消去
                DispInitial();

                // グリッド表示
                GridViewShow(dg);
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DispInitial();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // フォームを閉じます
            this.Close();
        }

        private string userFileSelect()
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = "社員、パートのCSVファイルを選択してください";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "ＣＳＶファイル(*.csv)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "CSVファイル確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFileSelect();
            if (sPath != string.Empty)
            {
                txtFileName.Text = sPath;
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
