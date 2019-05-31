using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using JMC.Common;

namespace JMC.prePrint
{
    public partial class prePrint : Form
    {
        string _csvFileName = string.Empty;

        const int PR_STATUS_MAIN = 1;   // 帳票区分：メイン勤務
        const int PR_STATUS_SUB = 2;    // 帳票区分：サブ勤務

        public prePrint(string fName, string sNum, int yakushokuType, string grpID, string grpName, int pMode)
        {
            InitializeComponent();

            _fName = fName;
            _sheetNum = sNum;
            _YakushokuType = yakushokuType;
            _grpID = grpID;
            _grpName = grpName;
            _pMode = pMode;
        }

        string _fName = string.Empty;   // ファイル名
        string _sheetNum = string.Empty;  // シート名
        string _grpID = string.Empty;   // グループID
        string _grpName = string.Empty;   // グループ名
        int _YakushokuType = 0;         // 表示社員の役職タイプ（１：パート、１以外：社員）
        int _pMode = 0;

        Utility.xlsShain[] xS = null;
        string[] csvArray = null;

        private void Form1_Load(object sender, EventArgs e)
        {
            // ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            // キャプション
            this.Text = "出勤簿プレ印刷 【" + _grpName + "】";

            // 自分自身のバージョン情報を取得する　2011/03/25
            System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(
                System.Reflection.Assembly.GetExecutingAssembly().Location);

            // キャプションにバージョンを追加　2011/03/25
            this.Text += " ver " + ver.FileMajorPart.ToString() + "." + ver.FileMinorPart.ToString();
            
            //// 部門コード・社員コード桁数取得
            //Utility.BumonShainKetasu.GetKetasu(_csvFileName);

            // 所属コード設定桁数
            //int ShozokuLen = global.ShozokuLength;
            int ShozokuLen = global.ShozokuMaxLength;

            // 社員コード範囲入力テキストボックス桁数
            //txtSNo.MaxLength = global.ShainLength;
            //txtENo.MaxLength = global.ShainLength;
            txtSNo.MaxLength = global.ShainMaxLength;
            txtENo.MaxLength = global.ShainMaxLength;

            switch (_pMode)
            {
                // 社員名簿エクセルシートモード
                case global.XLS_MODE:
                   
                    // 配列作成
                    Cursor = Cursors.WaitCursor;
                    Utility.ComboShain.xlsArrayLoad(_fName, _sheetNum, ref xS);
                    Cursor = Cursors.Default;

                    // 開始部門コンボロード
                    arrayCmbLoad(cmbBumonS, xS, global.flgOn);

                    // 終了部門コンボロード
                    arrayCmbLoad(cmbBumonE, xS, global.flgOn);

                    // 社員コンボロード
                    arrayCmbLoad(comboBox1, xS, global.flgOff);
                    //Utility.ComboShain.xlsArrayLoad(_fName, _sheetNum, comboBox1, global.flgOff);

                    // 特定部門コンボロード
                    arrayCmbLoad(comboBox2, xS, global.flgOn);
                    break;

                case global.CSV_MODE:
                    
                    // 社員名簿CSV読み込み
                    csvArray = System.IO.File.ReadAllLines(_fName, Encoding.Default);

                    // 開始部門コンボロード
                    Utility.ComboBumon.loadCsv(cmbBumonS, _fName);

                    // 終了部門コンボロード
                    Utility.ComboBumon.loadCsv(cmbBumonE, _fName);

                    // 社員コンボロード
                    Utility.ComboShain.loadCsv(comboBox1, _fName);

                    // 特定部門コンボロード
                    Utility.ComboBumon.loadCsv(comboBox2, _fName);
                    break;

                default:
                    break;
            }

            // DataGridViewの設定
            GridViewSetting(dg1);

            txtYear.Focus();

            // 社員番号
            txtSNo.Text = string.Empty;
            txtENo.Text = string.Empty;

            // チェックボタン
            btnCheckOn.Enabled = false;
            btnCheckOff.Enabled = false;

            // 印刷ボタン
            btnPrn.Enabled = false;

            // 元号表示　2011/03/24
            //label5.Text = Properties.Settings.Default.gengou;     // 2019/04/26 コメント化

            // 西暦化　2019/04/26
            label5.Text = "20";

            // 発行モード
            radioButton1.Checked = true;
            radioButton4.Checked = true;

            //// 有給記号非表示領域か判断する
            //_tBoxStatus = textBoxVisibleStatus(_dbName);
        }

        /// <summary>
        /// 選択したデータ領域が有休非表示領域か調べる
        /// </summary>
        /// <param name="dbName">PCA給与データ領域データベース名</param>
        /// <returns>true：有休非表示領域に該当、false：有休非表示領域に該当しない</returns>
        private bool textBoxVisibleStatus(string dbName)
        {
            bool rVal = true;

            // ローカルデータベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            OleDbDataReader dR;
            sCom.CommandText = "select * from 有休非表示領域 where Name = '" + dbName + "'";
            dR = sCom.ExecuteReader();
            rVal = dR.HasRows;
            dR.Close();
            sCom.Connection.Close();

            return rVal;
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///--------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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
                tempDGV.Height = 526;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(column);
                tempDGV.Columns.Add("col1", "コード");
                tempDGV.Columns.Add("col2", "所属");
                tempDGV.Columns.Add("col3", "社員番号");
                tempDGV.Columns.Add("col4", "社員名");
                tempDGV.Columns.Add("col5", "区分");

                tempDGV.Columns[0].Width = 40;
                tempDGV.Columns[1].Width = 70;
                tempDGV.Columns[2].Width = 300;
                tempDGV.Columns[3].Width = 80;
                tempDGV.Columns[4].Width = 200;
                tempDGV.Columns[5].Width = 80;

                tempDGV.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                tempDGV.ReadOnly = false;
                tempDGV.Columns[0].ReadOnly = false;
                tempDGV.Columns[1].ReadOnly = true;
                tempDGV.Columns[2].ReadOnly = true;
                tempDGV.Columns[3].ReadOnly = true;
                tempDGV.Columns[4].ReadOnly = true;
                tempDGV.Columns[5].ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
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

                //ソート機能制限
                for (int i = 0; i < tempDGV.Columns.Count; i++)
                {
                    tempDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ社員情報を表示する </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        /// <param name="sCode">
        ///     所属範囲開始コード</param>
        /// <param name="eCode">
        ///     所属範囲終了コード</param>
        /// <param name="sNo">
        ///     社員範囲開始コード</param>
        /// <param name="eNo">
        ///     社員範囲終了コード</param>
        ///--------------------------------------------------------------------------
        private void GridViewShowData(DataGridView tempDGV, int sCode, int eCode, string sNo, string eNo, string [] cArray)
        {
            try
            {
                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                System.Collections.ArrayList al = new System.Collections.ArrayList();

                foreach (var t in cArray)
                {
                    string[] f = t.Split(',');

                    if (f.Length < 4)
                    {
                        continue;
                    }

                    // 所属コード範囲指定
                    if (Utility.StrtoInt(f[3]) < sCode || Utility.StrtoInt(f[3]) > eCode)
                    {
                        continue;
                    }

                    // 社員番号範囲指定
                    int _sNo = Utility.StrtoInt(sNo);
                    int _eNo = 99999;

                    if (Utility.StrtoInt(eNo) != 0)
                    {
                        _eNo = Utility.StrtoInt(eNo);
                    }

                    if (Utility.StrtoInt(f[1]) < _sNo || Utility.StrtoInt(f[1]) > _eNo)
                    {
                        continue;
                    }

                    // 発行順
                    if (radioButton4.Checked)
                    {
                        // 社員番号順
                        al.Add(f[1] + "," + f[0] + "," + f[3] + "," + f[2] + "," + f[4]);
                    }
                    else
                    {
                        // 勤務先別、社員番号順
                        al.Add(f[3] + "," + f[2] + "," + f[1] + "," + f[0] + "," + f[4]);
                    }
                }

                al.Sort();

                foreach (var t in al)
                {
                    string [] f = t.ToString().Split(',');

                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();

                    tempDGV[0, iX].Value = true;

                    if (radioButton4.Checked)
                    {
                        // 社員番号順
                        tempDGV[1, iX].Value = f[2];
                        tempDGV[2, iX].Value = f[3];
                        tempDGV[3, iX].Value = f[0];
                        tempDGV[4, iX].Value = f[1];
                        tempDGV[5, iX].Value = f[4];
                    }
                    else
                    {
                        // 勤務先・社員番号順
                        tempDGV[1, iX].Value = f[0];
                        tempDGV[2, iX].Value = f[1];
                        tempDGV[3, iX].Value = f[2];
                        tempDGV[4, iX].Value = f[3];
                        tempDGV[5, iX].Value = f[4];
                    }

                    //tempDGV[5, iX].Value = string.Empty;

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

            //社員情報がないとき
            if (tempDGV.RowCount == 0)
            {
                MessageBox.Show("社員情報が存在しません", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                btnCheckOn.Enabled = false;
                btnCheckOff.Enabled = false;
                btnPrn.Enabled = false;
            }
            else
            {
                btnCheckOn.Enabled = true;
                btnCheckOff.Enabled = true;
                btnPrn.Enabled = true;
            }
        }


        ///--------------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ社員情報を表示する </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        /// <param name="sCode">
        ///     所属範囲開始コード</param>
        /// <param name="eCode">
        ///     所属範囲終了コード</param>
        /// <param name="sNo">
        ///     社員範囲開始コード</param>
        /// <param name="eNo">
        ///     社員範囲終了コード</param>
        ///--------------------------------------------------------------------------
        private void GridViewShowXls(DataGridView tempDGV, int sCode, int eCode, string sNo, string eNo, Utility.xlsShain [] cArray)
        {
            try
            {
                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                System.Collections.ArrayList al = new System.Collections.ArrayList();

                foreach (var t in cArray)
                {
                    //if (f.Length < 4)
                    //{
                    //    continue;
                    //}

                    // 所属コード範囲指定
                    if (t.bCode < sCode || t.bCode > eCode)
                    {
                        continue;
                    }

                    // 社員番号範囲指定
                    int _sNo = Utility.StrtoInt(sNo);
                    int _eNo = 99999;

                    if (Utility.StrtoInt(eNo) != 0)
                    {
                        _eNo = Utility.StrtoInt(eNo);
                    }

                    if (t.sCode < _sNo || t.sCode > _eNo)
                    {
                        continue;
                    }

                    // 発行順
                    if (radioButton4.Checked)
                    {
                        // 社員番号順
                        al.Add(t.sCode.ToString() + "," + t.sName + "," + t.bCode.ToString() + "," + t.bName);
                    }
                    else
                    {
                        // 勤務先別、社員番号順
                        al.Add(t.bCode.ToString() + "," + t.bName + "," + t.sCode.ToString() + "," + t.sName);
                    }
                }

                al.Sort();

                foreach (var t in al)
                {
                    string[] f = t.ToString().Split(',');

                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();

                    tempDGV[0, iX].Value = true;
                    if (radioButton4.Checked)
                    {
                        tempDGV[1, iX].Value = f[2].PadLeft(5, '0');
                        tempDGV[2, iX].Value = f[3];
                        tempDGV[3, iX].Value = f[0].PadLeft(5, '0');
                        tempDGV[4, iX].Value = f[1];
                    }
                    else
                    {
                        tempDGV[1, iX].Value = f[0].PadLeft(5, '0');
                        tempDGV[2, iX].Value = f[1];
                        tempDGV[3, iX].Value = f[2].PadLeft(5, '0');
                        tempDGV[4, iX].Value = f[3];
                    }

                    tempDGV[5, iX].Value = string.Format("{" + (_YakushokuType - 1) + "}", "社員", "パート");

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

            //社員情報がないとき
            if (tempDGV.RowCount == 0)
            {
                MessageBox.Show("社員情報が存在しません", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                btnCheckOn.Enabled = false;
                btnCheckOff.Enabled = false;
                btnPrn.Enabled = false;
            }
            else
            {
                btnCheckOn.Enabled = true;
                btnCheckOff.Enabled = true;
                btnPrn.Enabled = true;
            }
        }

        private void btnCheckOn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の社員全てを印刷対象にします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            for (int i = 0; i < dg1.Rows.Count; i++)
            {
                dg1[0, i].Value = true;
            }

            btnPrn.Enabled = true;
        }

        private void btnCheckOff_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の社員全てを印刷対象外にします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            for (int i = 0; i < dg1.Rows.Count; i++)
            {
                dg1[0, i].Value = false;
            }
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            int pCnt = 0;

            // エラーチェック
            if (ErrCheck() == false) return;

            // 件数取得
            pCnt = PrintRowCount();
            if (pCnt == 0)
            {
                MessageBox.Show("印刷対象行がありません", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show(pCnt.ToString() + "件の出勤簿を印刷します。よろしいですか？", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // 印刷
            sReport();

            MessageBox.Show("印刷が終了しました", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private Boolean ErrCheck()
        {
            if (Utility.NumericCheck(txtYear.Text) == false)
            {
                MessageBox.Show("年は数字で入力してください", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.NumericCheck(txtMonth.Text) == false)
            {
                MessageBox.Show("月は数字で入力してください", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("月が正しくありません", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }


            if (txtSNo.Text != string.Empty && Utility.NumericCheck(txtSNo.Text) == false)
            {
                MessageBox.Show("開始社員番号は数字で入力してください", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSNo.Focus();
                return false;
            }

            if (txtENo.Text != string.Empty && Utility.NumericCheck(txtENo.Text) == false)
            {
                MessageBox.Show("終了社員番号は数字で入力してください", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtENo.Focus();
                return false;
            }

            return true;
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     プリント件数取得 </summary>
        /// <returns>
        ///     印刷件数</returns>
        ///-------------------------------------------------------
        private int PrintRowCount()
        {
            int pCnt = 0;

            for (int i = 0; i < dg1.Rows.Count; i++)
            {
                if (dg1[0, i].Value.ToString() == "True") pCnt++;
            }

            return pCnt;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     出勤簿印刷・シート追加一括印刷 </summary>
        ///------------------------------------------------------------------
        private void sReport()
        {
            const int S_GYO = 7;       //エクセルファイル日付明細開始行

            //開始日付
            int StartDay = 1;

            //終了日付
            int EndDay = DateTime.DaysInMonth(int.Parse(txtYear.Text) + Utility.GetRekiHosei(), int.Parse(txtMonth.Text));  
            
            string sDate;
            DateTime eDate;

            //////const int S_ROWSMAX = 7; //エクセルファイル列最大値

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.OCR出勤簿シートパス, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // 休日配列インスタンス化
                    Config.Holiday[] Holiday = new Config.Holiday[1];
                    Holiday[0] = new Config.Holiday();
                    Holiday[0].hDate = DateTime.Parse("1900/01/01");

                    // ローカルデータベース接続
                    SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                    OleDbCommand sCom = new OleDbCommand();
                    sCom.Connection = Con.cnOpen();
                    OleDbDataReader dR;

                    // 休日データを配列に読み込む
                    int iH = 0;
                    string sqlSTRING = string.Empty;
                    sqlSTRING += "select * from 休日 order by 年月日";
                    sCom.CommandText = sqlSTRING;
                    dR = sCom.ExecuteReader();
                    while (dR.Read())
                    {
                        if (iH > 0)
                        {
                            Array.Resize(ref Holiday, iH + 1);  // 配列要素数追加
                            Holiday[iH] = new Config.Holiday(); // 休日配列インスタンス化
                        }

                        Holiday[iH].hDate = DateTime.Parse(dR["年月日"].ToString());

                        if (dR["月給者"].ToString() == "1") Holiday[iH].Gekkyuu = true;
                        else Holiday[iH].Gekkyuu = false;

                        if (dR["時給者"].ToString() == "1") Holiday[iH].Jikyuu = true;
                        else Holiday[iH].Jikyuu = false;

                        iH++;
                    }

                    dR.Close();
                    sCom.Connection.Close();

                    // ページカウント
                    int pCnt = 0;

                    //// 有給記号非表示設定 : 2017/11/27
                    //Excel.TextBox t34 = oxlsSheet.TextBoxes("Text Box 32");
                    //t34.Visible = false;
                    //oxlsSheet.Cells[5, 5] = string.Empty;

                    // パートのとき有給記号非表示設定 : 2017/11/30
                    if (_YakushokuType == global.STATUS_PART)
                    {
                        Excel.TextBox t34 = oxlsSheet.TextBoxes("Text Box 32");
                        t34.Visible = false;
                        oxlsSheet.Cells[5, 5] = string.Empty;
                    }


                    //Excel.TextBox t14 = oxlsSheet.TextBoxes("テキスト 14");

                    //t14.Visible = false;

                    //Excel.TextBox t34 = oxlsSheet.TextBoxes("yukyukigo");
                    //Excel.TextBox t14 = oxlsSheet.TextBoxes("kyukakigo");
                    //if (_tBoxStatus)
                    //{
                    //    t34.Visible = false;
                    //    t14.Visible = false;
                    //    oxlsSheet.Cells[5, 5] = string.Empty;
                    //}
                    //else
                    //{
                    //    t34.Visible = true;
                    //    t14.Visible = true;
                    //    oxlsSheet.Cells[5, 5] = "有給";
                    //}

                    // グリッドを順番に読む
                    for (int i = 0; i < dg1.RowCount; i++)
                    {
                        // チェックがあるものを対象とする
                        if (dg1[0, i].Value.ToString() == "True")
                        {
                            // 印刷2件目以降はシートを追加する
                            pCnt++;

                            if (pCnt > 1)
                            {
                                oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                                oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];
                            }

                            int sRow = i;

                            // 帳票区分（1:メイン勤務出勤簿、2:サブ勤務出勤簿）
                            // 配列から取得　2017/05/17
                            //if (radioButton1.Checked)
                            //{
                            //    oxlsSheet.Cells[2, 1] = PR_STATUS_MAIN.ToString();
                            //}
                            //else
                            //{
                            //    oxlsSheet.Cells[2, 1] = PR_STATUS_SUB.ToString();
                            //}

                            oxlsSheet.Cells[2, 1] = dg1[5, sRow].Value.ToString();

                            // 年
                            oxlsSheet.Cells[3, 4] = string.Format("{0, 2}", int.Parse(txtYear.Text)).Substring(0, 1);
                            oxlsSheet.Cells[3, 5] = string.Format("{0, 2}", int.Parse(txtYear.Text)).Substring(1, 1);

                            // 月
                            oxlsSheet.Cells[3, 8] = string.Format("{0, 2}", int.Parse(txtMonth.Text)).Substring(0, 1);
                            oxlsSheet.Cells[3, 9] = string.Format("{0, 2}", int.Parse(txtMonth.Text)).Substring(1, 1);

                            // 所属名
                            oxlsSheet.Cells[3, 15] = dg1[2, sRow].Value.ToString();

                            // 所属コード
                            string szCode = dg1[1, sRow].Value.ToString().PadLeft(global.ShozokuMaxLength, ' ');
                            for (int ci = 0; ci < szCode.Length; ci++)
                            {
                                oxlsSheet.Cells[3, 25 + ci] = szCode.Substring(ci, 1);
                            }

                            //oxlsSheet.Cells[3, 27] = dg1[1, sRow].Value.ToString().Substring(0, 1);
                            //oxlsSheet.Cells[3, 28] = dg1[1, sRow].Value.ToString().Substring(1, 1);
                            //oxlsSheet.Cells[3, 29] = dg1[1, sRow].Value.ToString().Substring(2, 1);

                            // 社員番号
                            //for (int ci = 0; ci < global.ShainLength; ci++)
                            //{
                            //    oxlsSheet.Cells[2, 25 + ci] = dg1[3, sRow].Value.ToString().Substring(ci, 1);
                            //}

                            oxlsSheet.Cells[2, 25] = dg1[3, sRow].Value.ToString().Substring(0, 1);
                            oxlsSheet.Cells[2, 26] = dg1[3, sRow].Value.ToString().Substring(1, 1);
                            oxlsSheet.Cells[2, 27] = dg1[3, sRow].Value.ToString().Substring(2, 1);
                            oxlsSheet.Cells[2, 28] = dg1[3, sRow].Value.ToString().Substring(3, 1);
                            oxlsSheet.Cells[2, 29] = dg1[3, sRow].Value.ToString().Substring(4, 1);

                            // 氏名
                            oxlsSheet.Cells[2, 15] = dg1[4, sRow].Value.ToString();

                            // 日付
                            int addRow = 0;
                            for (int iX = StartDay; iX <= EndDay; iX++)
                            {
                                // 暦補正値は設定ファイルから取得する
                                sDate = (int.Parse(txtYear.Text) + Utility.GetRekiHosei()).ToString() + "/" + txtMonth.Text + "/" + iX.ToString();
                                eDate = DateTime.Parse(sDate);
                                oxlsSheet.Cells[S_GYO + addRow, 2] = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);

                                rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO + addRow, 1];
                                rng[1] = (Excel.Range)oxlsSheet.Cells[S_GYO + addRow, 2];

                                // 日曜日なら曜日の背景色を変える
                                if (rng[1].Text.ToString() == "日")
                                {
                                    oxlsSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.LightGray;
                                }

                                // 祝祭日なら曜日の背景色を変える
                                for (int j = 0; j < Holiday.Length; j++)
                                {
                                    // 休日登録されている
                                    if (Holiday[j].hDate == eDate)
                                    {
                                        // 月給者または時給者が各々休日対象となっている
                                        if (dg1[5, sRow].Value.ToString() == "社員" && Holiday[j].Gekkyuu == true ||
                                            dg1[5, sRow].Value.ToString() == "パート" && Holiday[j].Jikyuu == true)
                                        {
                                            oxlsSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.LightGray;
                                            break;
                                        }
                                        else oxlsSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.White;
                                    }
                                }

                                // 行数加算
                                addRow++;
                            }
                        }
                    }

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    //oXls.Visible = true;

                    // 印刷
                    oXlsBook.PrintOut();

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();
                }

                finally
                {
                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            // マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();
            
            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtSNo) txtObj = txtSNo;
            if (sender == txtENo) txtObj = txtENo;

            txtObj.SelectAll();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("終了します。よろしいですか？",Application.ProductName,MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }
            this.Dispose();
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
            string sDate = string.Empty;
            int sCode;
            int eCode;
            string sNo;
            string eNo;

            //エラーチェック
            if (!ErrCheck()) return;
            
            //開始部門コード取得
            if (cmbBumonS.SelectedIndex == -1 || cmbBumonS.Text == string.Empty)
            {
                sCode = 0;
            }
            else
            {
                Utility.ComboBumon cmbs = new Utility.ComboBumon();
                cmbs = (Utility.ComboBumon)cmbBumonS.SelectedItem;
                sCode = Utility.StrtoInt(cmbs.code);
            }

            //終了部門コード取得
            if (cmbBumonE.SelectedIndex == -1 || cmbBumonE.Text == string.Empty)
            {
                string endingCode = "9";
                eCode = int.Parse(endingCode.PadLeft(global.ShozokuMaxLength, '9'));
                //eCode = 999;
            }
            else
            {
                Utility.ComboBumon cmbe = new Utility.ComboBumon();
                cmbe = (Utility.ComboBumon)cmbBumonE.SelectedItem;
                eCode = Utility.StrtoInt(cmbe.code);
            }

            //開始社員番号取得
            if (txtSNo.Text == string.Empty)
            {
                sNo = "00000";
            }
            else
            {
                sNo = txtSNo.Text.Trim().PadLeft(global.ShainLength, '0');
            }

            //終了社員番号取得
            if (txtENo.Text == string.Empty)
            {
                eNo = "99999";
            }
            else
            {
                eNo = txtENo.Text.Trim().PadLeft(global.ShainLength, '0');
            }

            //データ表示
            if (_pMode == global.XLS_MODE)
            {
                // エクセルモード
                GridViewShowXls(dg1, sCode, eCode, sNo, eNo, xS);
            }
            else if (_pMode == global.CSV_MODE)
            {
                // ＣＳＶモード
                GridViewShowData(dg1, sCode, eCode, sNo, eNo, csvArray);
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.KeyCode == Keys.Enter)
            //{
            //    if (!e.Control)
            //    {
            //        this.SelectNextControl(this.ActiveControl, !e.Shift, true, true, true);
            //    }
            //}
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {

            //if (e.KeyChar == (char)Keys.Enter)
            //{
            //    e.Handled = true;
            //}
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
 
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                label2.Enabled = true;
                cmbBumonS.Enabled = true;
                cmbBumonE.Enabled = true;
                label6.Enabled = true;

                label1.Enabled = true;
                txtSNo.Enabled = true;
                txtENo.Enabled = true;
                label7.Enabled = true;

                label8.Enabled = false;
                comboBox1.Enabled = false;
                label9.Enabled = false;
                comboBox2.Enabled = false;

                btnSel.Enabled = true;
                btnAdd.Enabled = false;
                btnPrn.Enabled = false;

                dg1.RowCount = 0;

                groupBox2.Enabled = true;
                radioButton4.Enabled = true;
                radioButton3.Enabled = true;
            }
            else
            {
                label2.Enabled = false;
                cmbBumonS.Enabled = false;
                cmbBumonE.Enabled = false;
                label6.Enabled = false;

                label1.Enabled = false;
                txtSNo.Enabled = false;
                txtENo.Enabled = false;
                label7.Enabled = false;

                label8.Enabled = true;
                comboBox1.Enabled = true;
                label9.Enabled = true;
                comboBox2.Enabled = true;

                btnSel.Enabled = false;
                btnAdd.Enabled = true;
                btnPrn.Enabled = true;

                //チェックボタン
                btnCheckOn.Enabled = false;
                btnCheckOff.Enabled = false;

                dg1.RowCount = 0;

                groupBox2.Enabled = false;
                radioButton4.Enabled = false;
                radioButton3.Enabled = false;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // 個別印刷対象者情報をグリッドビューへ追加します
            gridViewAdd();
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     個別印刷対象者情報をグリッドビューへ追加します </summary>
        ///-----------------------------------------------------------
        private void gridViewAdd()
        {
            // 社員未選択
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("印刷する社員を選択して下さい", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                comboBox1.Focus();
                return;
            }

            // 所属未選択
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("所属を選択して下さい", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                comboBox2.Focus();
                return;
            }

            // グリッドへ追加
            dg1.Rows.Add();
            dg1[0, dg1.Rows.Count - 1].Value = "True";

            // 部門コンボ情報
            Utility.ComboBumon cmbB = (Utility.ComboBumon)comboBox2.SelectedItem;
            dg1[1, dg1.Rows.Count - 1].Value = cmbB.code;
            dg1[2, dg1.Rows.Count - 1].Value = cmbB.Name;

            // 社員コンボ情報
            Utility.ComboShain cmbS = (Utility.ComboShain)comboBox1.SelectedItem;
            dg1[3, dg1.Rows.Count - 1].Value = cmbS.code;
            dg1[4, dg1.Rows.Count - 1].Value = cmbS.Name;
            dg1[5, dg1.Rows.Count - 1].Value = "";
       
            // コンボボックスクリア
            comboBox1.SelectedIndex = -1;
            comboBox1.Text = string.Empty;
            comboBox2.SelectedIndex = -1;
            comboBox2.Text = string.Empty;
        }



        ///----------------------------------------------------------------
        /// <summary>
        ///     配列からコンボボックスにロードする </summary>
        /// <param name="tempObj">
        ///     コンボボックスオブジェクト</param>
        /// <param name="fName">
        ///     ＣＳＶデータファイルパス</param>
        /// <param name="cmbKbn">
        ///     ０：社員コンボボックス、１：部門コンボボックス</param>
        ///----------------------------------------------------------------
        private static void arrayCmbLoad(ComboBox tempObj, Utility.xlsShain [] x, int cmbKbn)
        {
            string tl = "";

            try
            {
                if (cmbKbn == global.flgOff)
                {
                    tl = "社員";
                }
                else
                {
                    tl = "部門";
                }

                Utility.ComboBumon cmb1;    // 部門コンボボックス
                Utility.ComboShain cmbS;    // 社員コンボボックス

                tempObj.Items.Clear();
                tempObj.DisplayMember = "DisplayName";
                tempObj.ValueMember = "code";

                // 社員名簿配列読み込み
                System.Collections.ArrayList al = new System.Collections.ArrayList();
                string bn = "";

                foreach (var t in x)
                {
                    if (cmbKbn == global.flgOff)
                    {
                        // 社員
                        bn = t.sCode.ToString().PadLeft(5, '0') + "," + t.sName + "";
                    }
                    else
                    {
                        // 部門
                        bn = t.bCode.ToString().PadLeft(5, '0') + "," + t.bName + "";
                    }

                    al.Add(bn);
                }

                // 配列をソートします
                al.Sort();

                string alCode = string.Empty;

                foreach (var item in al)
                {
                    string[] d = item.ToString().Split(',');

                    // 重複はネグる
                    if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
                    {
                        continue;
                    }

                    // コンボボックスにセット
                    if (cmbKbn == global.flgOn)
                    {
                        // 部門コンボボックス
                        cmb1 = new Utility.ComboBumon();
                        cmb1.ID = 0;
                        cmb1.DisplayName = item.ToString().Replace(',', ' ');

                        string[] cn = item.ToString().Split(',');
                        cmb1.Name = cn[1] + "";
                        cmb1.code = cn[0] + "";
                        tempObj.Items.Add(cmb1);
                    } 
                    else if (cmbKbn == global.flgOff)
                    {
                        // 社員コンボボックス
                        cmbS = new Utility.ComboShain();
                        cmbS.ID = 0;
                        cmbS.DisplayName = item.ToString().Replace(',', ' ');

                        string[] cn = item.ToString().Split(',');
                        cmbS.Name = cn[1] + "";
                        cmbS.code = cn[0] + "";
                        tempObj.Items.Add(cmbS);
                    }

                    alCode = item.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,  tl + "コンボボックスロード");
            }
        }
    }
}
