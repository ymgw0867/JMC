using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Odbc;
using JMC.Common;

namespace JMC
{
    public partial class frmComSelect : Form
    {
        public frmComSelect(int cMode)
        {
            InitializeComponent();
            _cMode = cMode;
            adp.Fill(dts.社員ファイル);

        }

        int _cMode = 0;
        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.社員ファイルTableAdapter adp = new DataSet1TableAdapters.社員ファイルTableAdapter();

        private void frmComSelect_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            // DataGridViewの設定
            GridViewSetting(dg1);

            // データ表示
            GridViewShainFile(dg1);

            // 終了時タグ初期化
            Tag = string.Empty;

            // モード
            label1.Enabled = true;
            txtFileName.Enabled = true;
            btnSel.Enabled = true;

            if (_cMode == 1)
            {
                label1.Enabled = false;
                txtFileName.Enabled = false;
                btnSel.Enabled = false;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います　</summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 224;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col6", "グループ名");
                tempDGV.Columns.Add("col1", "区分");
                tempDGV.Columns.Add("col2", "ファイル名");
                //tempDGV.Columns.Add("col3", "シート名");
                tempDGV.Columns.Add("col4", "");
                tempDGV.Columns.Add("col5", "");

                tempDGV.Columns[0].Width = 200;
                tempDGV.Columns[1].Width = 70;
                ////tempDGV.Columns[1].Width = 200;
                //tempDGV.Columns[3].Width = 140;
                tempDGV.Columns[3].Visible = false;
                tempDGV.Columns[4].Visible = false;

                tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                //tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = false;

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

        ///---------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ会社情報を表示する </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>       
        ///---------------------------------------------------------------
        private void GridViewShowData(DataGridView tempDGV)
        {
            dbControl.DataControl dcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
            OleDbDataReader dR = null;

            try
            {
                // データリーダー取得
                string mySql = string.Empty;
                mySql += "SELECT * FROM Common_Unit_DataAreaInfo ";
                mySql += "where CompanyTerm = " + DateTime.Today.Year.ToString();
                dR = dcon.FreeReader(mySql);

                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    // "CompanyCode"が数字のレコードを対象とする
                    if (Utility.NumericCheck(dR["CompanyCode"].ToString()))
                    {
                        //データグリッドにデータを表示する
                        tempDGV.Rows.Add();
                        tempDGV[1, iX].Value = dR["CompanyCode"].ToString();        // コード
                        tempDGV[2, iX].Value = dR["CompanyName"].ToString().Trim(); // 会社名
                        tempDGV[3, iX].Value = dR["CompanyTerm"].ToString().Trim(); // 年度
                        tempDGV[4, iX].Value = dR["Name"].ToString().Trim();        // データベース名
                        iX++;
                    }
                }
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                dcon.Close();
            }

            //会社情報がないとき
            if (tempDGV.RowCount == 0) 
            {
                MessageBox.Show("会社領域情報が存在しません", "会社領域選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ会社情報を表示する </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>       
        ///---------------------------------------------------------------
        private void GridViewShainFile(DataGridView tempDGV)
        {
            try
            {
                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                foreach (var t in dts.社員ファイル.Where(a => a.ID > 0).OrderBy(a => a.ID))
                {
                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();

                    tempDGV[0, iX].Value = t.グループ名;

                    if (t.区分 == global.STATUS_SHAIN)
                    {
                        tempDGV[1, iX].Value = "社員";
                    }
                    else if (t.区分 == global.STATUS_PART)
                    {
                        tempDGV[1, iX].Value = "パート";
                    }
                    else
                    {
                        tempDGV[1, iX].Value = "";
                    }

                    tempDGV[2, iX].Value = t.ファイル名;
                    //tempDGV[3, iX].Value = t.シート名.ToString();
                    tempDGV[3, iX].Value = t.ID.ToString();
                    tempDGV[4, iX].Value = t.区分.ToString();
                    iX++;
                }

                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
            }

            // 社員情報がないとき
            if (tempDGV.RowCount == 0)
            {
                MessageBox.Show("社員マスタ情報が存在しません", "社員ファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // 社員ファイル情報がないときはそのままクローズ
            if (dg1.RowCount == 0)
            {
                _pFileName = string.Empty;
            }
            else
            {
                if (dg1.SelectedRows.Count == 0 && txtFileName.Text == string.Empty)
                {
                    MessageBox.Show("ファイルを選択してください", "未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (dg1.SelectedRows.Count != 0)
                {
                    // エクセルシートモード
                    // グループ名を取得する
                    _PName = dg1[0, dg1.SelectedRows[0].Index].Value.ToString();

                    // グループIDを取得する
                    _pID = dg1[3, dg1.SelectedRows[0].Index].Value.ToString();

                    // 名簿ファイル名を取得する
                    _pFileName = dg1[2, dg1.SelectedRows[0].Index].Value.ToString();

                    //// シート名を取得する
                    //_pSheetNum = dg1[3, dg1.SelectedRows[0].Index].Value.ToString();

                    // 社員・パート区分を取得する
                    _pYakushokuType = Utility.StrtoInt(dg1[4, dg1.SelectedRows[0].Index].Value.ToString());
                }
                else
                {
                    // ＣＳＶモード
                    // グループIDを取得する
                    _pID = "";
                    _PName = "ＣＳＶファイル";

                    // 名簿ファイル名を取得する
                    _pFileName = txtFileName.Text;

                    // シート名を取得する
                    _pSheetNum = "";

                    // 社員・パート区分を取得する
                    _pYakushokuType = 0;
                }

                //フォームを閉じる
                Tag = "btn";
                this.Close();
            }
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     出勤簿社員名簿ファイルパス取得 </summary>
        ///-------------------------------------------------------------
        private string  getcsvName()
        {
            if (txtFileName.Text == string.Empty)
            {
                MessageBox.Show("社員ファイルを選択してください", "ファイル未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return string.Empty;
            }

            if (!System.IO.File.Exists(txtFileName.Text))
            {
                MessageBox.Show("指定されたファイルは存在しません", "ファイルパス確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return string.Empty;
            }

            // 名簿ファイル名を返す
            return txtFileName.Text;
        }


        private void frmComSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Tag.ToString() == string.Empty)
                {
                    if (MessageBox.Show("プログラムを終了します。よろしいですか？", "終了", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        //終了処理
                        Environment.Exit(0);
                    }
                    else
                    {
                        e.Cancel = true;
                        return;
                    }
                }
            }
        }

        //// 社員ファイル名
        //public string _ComDataBaseName { get; set; }

        //// 会社領域名
        //public string _ComName { get; set; }

        // グループID
        public string _pID { get; set; }

        // グル－プ名
        public string _PName { get; set; }

        // 出勤簿名簿CSVファイル名
        public string _pFileName { get; set; }

        // シート番号
        public string _pSheetNum { get; set; }

        // 社員・パート区分
        public int _pYakushokuType { get; set; }

        private void btnSel_Click(object sender, EventArgs e)
        {
            string tl = "出勤簿用名簿CSVデータファイル選択";
            string fl = "CSVファイル(*.CSV)|*.csv|全てのファイル(*.*)|*.*";

            // 名簿ファイルを選択する
            string sPath = Utility.userFileSelect(tl, fl);
            if (sPath != string.Empty)
            {
                txtFileName.Text = sPath;

                // 給与グループの選択状態を解除する
                dg1.ClearSelection();
            }
        }
    }
}
