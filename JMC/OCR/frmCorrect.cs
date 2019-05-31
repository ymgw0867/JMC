using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using JMC.Common;

namespace JMC.OCR
{
    public partial class frmCorrect : Form
    {
        ///-------------------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ<summary>
        /// <param name="fName">
        ///     社員ファイル名</param>
        /// <param name="sNum">
        ///     シート番号</param>
        /// <param name="yakushokuType">
        ///     社員・パート区分</param>
        /// <param name="sID">
        ///     処理モード</param>
        ///-------------------------------------------------------------------------
        public frmCorrect(string fName, string sNum, int yakushokuType, string grpID, string sID)
        {
            InitializeComponent();
            dID = sID;              // 処理モード
            _fName = fName;     // ファイル名
            _sheetNum = sNum;     // シート名
            _YakushokuType = yakushokuType; // 社員・パート区分
            _grpID = grpID; // グループID

            // 休日配列
            Holiday[0] = new Config.Holiday();
            Holiday[0].hDate = DateTime.Parse("1900/01/01");

            // 給与グループファイル読み込み
            adp.Fill(dts.社員ファイル);
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.社員ファイルTableAdapter adp = new DataSet1TableAdapters.社員ファイルTableAdapter();

        //datagridview表示行数
        const int _MULTIGYO = 31;

        // MDBデータキー配列
        string[] sID;

        //カレントデータインデックス
        int cI;

        // 社員マスターより取得した所属コード
        string mSzCode = string.Empty;

        //終了ステータス
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";

        //bool bDrag = false;
        //Point posStart;

        string dID = string.Empty;  // 表示する過去データのID
        string _fName = string.Empty;   // ファイル名
        string _sheetNum = string.Empty;  // シート名
        string _grpID = string.Empty;   // グループ名
        Config.Holiday[] Holiday = new Config.Holiday[1];   // 休日配列インスタンス化
        int _YakushokuType = 0;                     // 表示社員の役職タイプ（１：パート、１以外：社員）
        int _ShainID = 0;                           // 社員ＩＤ

        // パート変形労働時間制における労働時間総枠の配列
        double[,] ptLimitTm = { { 31, 177.0 }, { 30, 171.0 }, { 29, 165.0 }, { 28, 160.0 } };
        //double[,] ptLimitTm = { { 31, 177.1 }, { 30, 171.4 }, { 29, 165.7 }, { 28, 160.0 } }; 2013/07/05 小数点以下切り捨て

        // 社員マスター配列 2016/11/15
        Utility.xlsShain[] xS = null;

        // グループフォルダ
        string _gDir = string.Empty;

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridviewSet.Setting(dataGridView1);

            //// 部門コード・社員コード桁数取得 ：2013/07/12 勤務データ作成時のみ
            //if (dID == string.Empty) Utility.BumonShainKetasu.GetKetasu(_PCADBName);

            //元号を取得
            label1.Text = Properties.Settings.Default.gengou;

            // 休日配列作成
            GetHolidayArray();

            // 勤務データ登録
            if (dID == string.Empty)
            {
                // エクセル社員情報配列読み込み
                //Utility.ComboShain.xlsArrayLoad(_fName, _sheetNum, ref xS);
                Utility.ComboShain.csvArrayLoad(_fName, ref xS);

                //CSVデータをMDBへ読み込む
                GetCsvDataToMDB();

                //MDB件数カウント
                if (CountMDB() == 0)
                {
                    MessageBox.Show("対象となる出勤簿データがありません", "出勤簿データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Environment.Exit(0);     //終了処理
                }

                //MDBデータキー項目読み込み
                sID = LoadMdbID();

                //エラー情報初期化
                ErrInitial();

                //最初のレコードを表示
                cI = 0;
                DataShow(cI, sID, this.dataGridView1);

                //キャプション
                this.Text = "出勤簿データ登録";

                // 給与グループフォルダを作成 : 2016/12/07
                createGrpDir(_grpID);
            }
            else
            {
                //キャプション
                this.Text = "過去出勤簿データ表示 ";

                // 過去データの表示
                LastDataShow(dID, dataGridView1);
            }

            //tagを初期化
            this.Tag = string.Empty;
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     グループフォルダの生成 </summary>
        /// <param name="g">
        ///     グループID</param>
        ///--------------------------------------------------------
        private void createGrpDir(string g)
        {
            _gDir = g.PadLeft(3, '0');

            if (dts.社員ファイル.Any(a => a.ID == Utility.StrtoInt(g)))
            {
                // グループ名を取得してフォルダ名を決定する
                var s = dts.社員ファイル.Single(a => a.ID == Utility.StrtoInt(g));

                if (!s.Isグループ名Null())
                {
                    _gDir += " " + s.グループ名;
                }

                // 未作成ならフォルダを作成
                if (!Directory.Exists(Properties.Settings.Default.instPath + _gDir))
                {
                    Directory.CreateDirectory(Properties.Settings.Default.instPath + _gDir);
                }
            }
        }

        // カラム定義
        private static string cDay = "col1";
        private static string cWeek = "col2";
        private static string cKyuka = "colKyuka";
        private static string cYukyu = "col3";
        private static string cSH = "col4";
        private static string cSE = "col16";
        private static string cSM = "col5";
        private static string cEH = "col6";
        private static string cEE = "col17";
        private static string cEM = "col7";
        private static string cKKH = "col8";
        private static string cKKE = "col18";
        private static string cKKM = "col9";
        private static string cKSH = "col10";
        private static string cKSE = "col19";
        private static string cKSM = "col11";
        private static string cTH = "col12";
        private static string cTE = "col20";
        private static string cTM = "col13";
        private static string cCheck = "col14";
        private static string cID = "col15";

        // データグリッドビュークラス
        private class GridviewSet
        {
            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
            public static void Setting(DataGridView tempDGV)
            {
                try
                {
                    //フォームサイズ定義

                    // 列スタイルを変更する

                    tempDGV.EnableHeadersVisualStyles = false;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (Single)9.5, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                    tempDGV.ColumnHeadersHeight = 22;
                    tempDGV.RowTemplate.Height = 22;

                    // 全体の高さ
                    tempDGV.Height = 706;
                    // 全体の幅
                    tempDGV.Width = 480;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                    //各列幅指定
                    tempDGV.Columns.Add(cDay, "日");
                    tempDGV.Columns.Add(cWeek, "曜");
                    tempDGV.Columns.Add(cKyuka, "休暇");
                    tempDGV.Columns.Add(cYukyu, "有給");
                    tempDGV.Columns.Add(cSH, "開");
                    tempDGV.Columns.Add(cSE, "");
                    tempDGV.Columns.Add(cSM, "始");
                    tempDGV.Columns.Add(cEH, "終");
                    tempDGV.Columns.Add(cEE, "");
                    tempDGV.Columns.Add(cEM, "了");
                    tempDGV.Columns.Add(cKKH, "規");
                    tempDGV.Columns.Add(cKKE, "");
                    tempDGV.Columns.Add(cKKM, "定");
                    tempDGV.Columns.Add(cKSH, "深");
                    tempDGV.Columns.Add(cKSE, "");
                    tempDGV.Columns.Add(cKSM, "夜");
                    tempDGV.Columns.Add(cTH, "実");
                    tempDGV.Columns.Add(cTE, "");
                    tempDGV.Columns.Add(cTM, "働");

                    DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                    tempDGV.Columns.Add(column);
                    tempDGV.Columns[19].Name = cCheck;
                    tempDGV.Columns[19].HeaderText = "";

                    tempDGV.Columns.Add(cID, "");   // 明細ID
                    tempDGV.Columns[cID].Visible = false;

                    foreach (DataGridViewColumn c in tempDGV.Columns)
                    {
                        // 幅                       
                        if (c.Name == cSE || c.Name == cEE || c.Name == cKKE ||
                            c.Name == cKSE || c.Name == cTE) c.Width = 10;
                        else c.Width = 28;

                        tempDGV.Columns[cCheck].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        
                        // 表示位置
                        if (c.Index < 4 || c.Name == cSE || c.Name == cEE || c.Name == cKKE || 
                            c.Name == cKSE || c.Name == cTE) 
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        else c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        if (c.Name == cSH || c.Name == cEH || c.Name == cKKH || 
                            c.Name == cKSH || c.Name == cTH) 
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        if (c.Name == cSM || c.Name == cEM || c.Name == cKKM || 
                            c.Name == cKSM || c.Name == cTM)
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;

                        if (c.Name == cCheck) c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                        // 編集可否
                        if (c.Index < 2 || c.Name == cSE || c.Name == cEE || c.Name == cKKE ||
                            c.Name == cKSE || c.Name == cTE) 
                            c.ReadOnly = true;
                        else c.ReadOnly = false;

                        // 区切り文字
                        if (c.Name == cSE || c.Name == cEE || c.Name == cKKE || c.Name == cKSE || c.Name == cTE) 
                            c.DefaultCellStyle.Font = new Font("ＭＳＰゴシック", 8, FontStyle.Regular);
 
                    }

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // 編集可とする
                    //tempDGV.ReadOnly = false;

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

                    // ソート禁止
                    foreach (DataGridViewColumn c in tempDGV.Columns)
                    {
                        c.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

                    // 編集モード
                    tempDGV.EditMode = DataGridViewEditMode.EditOnEnter;

                    // 入力可能桁数
                    foreach (DataGridViewColumn c in tempDGV.Columns)
                    {
                        if (c.Name != cCheck)
                        {
                            DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;
                            if (c.Name == cKyuka || c.Name == cYukyu || 
                                c.Name == cKKH || c.Name == cKSH) col.MaxInputLength = 1;
                            else col.MaxInputLength = 2;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする </summary>
        ///----------------------------------------------------------
        private void GetCsvDataToMDB()
        {
            //CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv");

            //CSVファイルがなければ終了
            int cTotal = 0;
            if (inCsv.Length == 0) return;
            else cTotal = inCsv.Length;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // データベースへ接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    // 行カウント
                    int sline = 0;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = cCnt / cTotal * 100;
                    frmP.ProgressStep();

                    // OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    string fn = Path.GetFileName(files);
                    if (fn.Length == global.CSVFILENAMELENGTH)
                    {
                        StringBuilder sb = new StringBuilder();

                        // ヘッダID
                        string hdID = string.Empty;

                        // CSVファイルインポート
                        var s = File.ReadAllLines(files, Encoding.Default);
                        foreach (var stBuffer in s)
                        {
                            if (stBuffer == string.Empty) break;

                            // カンマ区切りで分割して配列に格納する
                            string[] stCSV = stBuffer.Split(',');

                            // MDBへ登録する
                            // 勤務記録ヘッダテーブル
                            if (sline == 0)
                            {
                                sb.Clear();
                                sb.Append("insert into 出勤簿ヘッダ ");
                                sb.Append("(ID,社員ID,個人番号,氏名,年,月,所属コード,給与区分,画像名,出勤日数合計,");
                                sb.Append("有休日数合計,有休時間合計,特休日数合計,振休日数合計,振出日数合計,遅刻早退回数,欠勤日数合計,実稼動日数合計,");
                                sb.Append("総労働,総労働分,残業時,残業分,深夜勤務時間合計,月間規定勤務時間,");
                                sb.Append("パート労働時間総枠,確認,データ領域名,更新年月日,立替金,旅費交通費,勤務先区分) ");
                                sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

                                sCom.CommandText = sb.ToString();
                                sCom.Parameters.Clear();

                                // ID
                                if (stCSV[0].Length > 17)
                                {
                                    sCom.Parameters.AddWithValue("@ID", stCSV[0].Substring(0, 17));
                                }
                                else
                                {
                                    sCom.Parameters.AddWithValue("@ID", stCSV[0]);
                                }

                                hdID = stCSV[0];


                                //// PCA給与より社員情報を取得します　SQLServer接続 ////////////////////////////////////////////////////////////////////
                                //dbControl.DataControl dCon = new dbControl.DataControl(_PCADBName);
                                //OleDbDataReader dR;

                                //// 社員情報データリーダーを取得する
                                //StringBuilder sb2 = new StringBuilder();
                                //sb2.Append("select * from Shain ");
                                //sb2.Append("where Shain.Code = '" + stCSV[5].Trim().PadLeft(5, '0') + "'");

                                //dR = dCon.FreeReader(sb2.ToString());

                                //string sName = string.Empty;
                                //string sYakushokuType = "0";
                                //string sID = "0";

                                //while (dR.Read())
                                //{
                                //    sName = dR["Sei"].ToString().Trim() + " " + dR["Mei"].ToString().Trim(); // 氏名取得
                                //    sYakushokuType = dR["YakushokuType"].ToString();    // 役職タイプ取得
                                //    sID = dR["Id"].ToString();                     // 社員ＩＤ
                                //}

                                //dR.Close();
                                //dCon.Close();
                                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                                // 社員ID
                                string pID = global.flgOff.ToString();    // 2016/11/16　不使用のため全てゼロとする
                                sCom.Parameters.AddWithValue("@sID", pID);

                                // 個人番号
                                if (stCSV[5].Length > global.ShainLength)
                                {
                                    sCom.Parameters.AddWithValue("@kjn", stCSV[5].Substring(0, global.ShainLength));
                                }
                                else
                                {
                                    sCom.Parameters.AddWithValue("@kjn", stCSV[5].Trim().PadLeft(5, '0'));
                                }

                                // 氏名
                                string zCode = string.Empty;
                                string zName = string.Empty;
                                string sName = Utility.ComboShain.getXlsSname(xS, Utility.StrtoInt(stCSV[5].Trim().PadLeft(5, '0')), out zCode, out zName);
                                sCom.Parameters.AddWithValue("@Name", sName);

                                // 年
                                if (stCSV[2].Length > 2)
                                {
                                    sCom.Parameters.AddWithValue("@year", stCSV[3].Substring(0, 2));
                                }
                                else
                                {
                                    sCom.Parameters.AddWithValue("@year", stCSV[3]);
                                }

                                // 月
                                string m = string.Empty;

                                // 2桁以内の文字列で取得
                                if (stCSV[4].Length > 2)
                                {
                                    m = stCSV[4].Substring(0, 2);
                                }
                                else
                                {
                                    m = stCSV[4];
                                }

                                // 数字以外はempty
                                if (!Utility.NumericCheck(m))
                                {
                                    m = string.Empty;
                                }
                                else
                                {
                                    m = int.Parse(m).ToString();
                                }
                                sCom.Parameters.AddWithValue("@month", m);

                                // 所属コード
                                if (stCSV[6].Length > global.ShozokuLength)
                                {
                                    sCom.Parameters.AddWithValue("@Szk", stCSV[6].Substring(0, global.ShozokuLength));
                                }
                                else
                                {
                                    sCom.Parameters.AddWithValue("@Szk", stCSV[6]);
                                }
                                
                                sCom.Parameters.AddWithValue("@kyuk", _YakushokuType);        // 給与区分
                                sCom.Parameters.AddWithValue("@IMG", Utility.getStringSubMax(stCSV[1], 21));    // 画像名
                                sCom.Parameters.AddWithValue("@t1", "0");           // 出勤日数合計
                                sCom.Parameters.AddWithValue("@t2", "0");           // 有休日数合計
                                sCom.Parameters.AddWithValue("@t3", "0");           // 有休時間合計
                                sCom.Parameters.AddWithValue("@t4", "0");           // 特休日数合計
                                sCom.Parameters.AddWithValue("@furi", "0");         // 振休日数合計
                                sCom.Parameters.AddWithValue("@furide", "0");       // 振出日数合計
                                sCom.Parameters.AddWithValue("@chi", "0");          // 遅刻早退回数
                                sCom.Parameters.AddWithValue("@t5", "0");           // 欠勤日数合計
                                sCom.Parameters.AddWithValue("@t6", "0");           // 実稼動日数合計
                                sCom.Parameters.AddWithValue("@t7", "0");           // 総労働
                                sCom.Parameters.AddWithValue("@t8", "0");           // 総労働分
                                sCom.Parameters.AddWithValue("@t9", "0");           // 残業時
                                sCom.Parameters.AddWithValue("@t10", "0");          // 残業分
                                sCom.Parameters.AddWithValue("@t11", "0");          // 深夜勤務時間合計
                                sCom.Parameters.AddWithValue("@t12", "0");          // 月間規定勤務時間
                                sCom.Parameters.AddWithValue("@t13", "0");          // パート労働時間総枠
                                sCom.Parameters.AddWithValue("@t14", "0");          // 確認
                                sCom.Parameters.AddWithValue("@t15", _grpID);       // データ領域名
                                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());  // 更新年月日
                                sCom.Parameters.AddWithValue("@t201611_1", "0");    // 立替金
                                sCom.Parameters.AddWithValue("@t201611_2", "0");    // 旅費交通費
                                sCom.Parameters.AddWithValue("@subMain", Utility.getStringSubMax(stCSV[2], 1));     // 勤務先区分 : 2018/03/31

                                // テーブル書き込み
                                sCom.ExecuteNonQuery();
                            }
                            else
                            {
                                // 出勤簿明細テーブル
                                sb.Clear();
                                sb.Append("insert into 出勤簿明細 ");
                                sb.Append("(ヘッダID,日付,休暇記号,有給記号,開始時,開始分,終了時,終了分,規定内時,");
                                sb.Append("規定内分,深夜帯時,深夜帯分,実働時,実働分,更新年月日) ");
                                sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                                sCom.CommandText = sb.ToString();

                                sCom.Parameters.Clear();
                                sCom.Parameters.AddWithValue("@HDID", hdID);        // ヘッダID
                                sCom.Parameters.AddWithValue("@day", stCSV[0]);     // 日付

                                // getStringSubMaxを使用 : 2018/03/31
                                sCom.Parameters.AddWithValue("@Kyuka", Utility.getStringSubMax(stCSV[1], 1));   // 休暇記号
                                sCom.Parameters.AddWithValue("@yukyu", Utility.getStringSubMax(stCSV[2], 1));   // 有給休暇
                                sCom.Parameters.AddWithValue("@sh", Utility.getStringSubMax(stCSV[3], 2));      // 開始時
                                sCom.Parameters.AddWithValue("@sm", Utility.getStringSubMax(stCSV[4], 2));      // 開始分
                                sCom.Parameters.AddWithValue("@eh", Utility.getStringSubMax(stCSV[5], 2));      // 終了時
                                sCom.Parameters.AddWithValue("@em", Utility.getStringSubMax(stCSV[6], 2));      // 終了分
                                sCom.Parameters.AddWithValue("@kh", Utility.getStringSubMax(stCSV[7], 1));      // 規定内時
                                sCom.Parameters.AddWithValue("@km", Utility.getStringSubMax(stCSV[8], 2));      // 規定内分
                                sCom.Parameters.AddWithValue("@ksh", Utility.getStringSubMax(stCSV[9], 1));     // 深夜帯時
                                sCom.Parameters.AddWithValue("@ksm", Utility.getStringSubMax(stCSV[10], 2));    // 深夜帯分
                                sCom.Parameters.AddWithValue("@th", Utility.getStringSubMax(stCSV[11], 2));     // 実働時
                                sCom.Parameters.AddWithValue("@tm", Utility.getStringSubMax(stCSV[12], 2));     // 実働分
                                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());      // 更新年月日

                                // テーブル書き込み
                                sCom.ExecuteNonQuery();                                
                            }

                            // 行カウントインクルメント
                            sline++;
                        }
                    }
                }

                // トランザクションコミット
                sTran.Commit();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "出勤簿CSVインポート処理", MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }

        }

        /// <summary>
        /// 出勤簿ヘッダデータの件数をカウントする
        /// </summary>
        /// <returns>データ件数</returns>
        private int CountMDB()
        {
            int rCnt = 0;
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;
            mySql += "select ID from 出勤簿ヘッダ order by ID";

            sCom.CommandText = mySql;
            sCom.Connection = Con.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }

        /// <summary>
        /// MDB明細データの件数をカウントする
        /// </summary>
        /// <returns>レコード件数</returns>
        private int CountMDBitem()
        {
            int rCnt = 0;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;

            mySql += "select ID from 出勤簿明細 order by ID";

            sCom.CommandText = mySql;
            sCom.Connection = dCon.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     MDBデータのキー項目を配列に読み込む </summary>
        /// <returns>
        ///     キー配列</returns>
        ///----------------------------------------------------------
        private string[] LoadMdbID()
        {
            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB();
            string [] DenID = new string[1];
            int rCnt = 1;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;

            //mySql += "select ID from 出勤簿ヘッダ order by 個人番号,勤務先区分,ID"; // 2016/11/17
            mySql += "select ID from 出勤簿ヘッダ order by ID"; // 2017/05/08

            sCom.CommandText = mySql;
            sCom.Connection = dCon.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //プログレスバー表示
                frmP.Text = "出勤簿データロード中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt / cTotal * 100;
                frmP.ProgressStep();

                //2件目以降は要素数を追加
                //if (rCnt > 1) DenID.CopyTo(DenID = new string[rCnt], 0);
                if (rCnt > 1) Array.Resize(ref DenID, rCnt);
                DenID[rCnt - 1] = dR["ID"].ToString();

                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return DenID;
        }

        private void ErrInitial()
        {
            //エラー情報初期化
            lblErrMsg.Visible = false;
            global.errNumber = global.eNothing;     //エラー番号
            global.errMsg = string.Empty;           //エラーメッセージ
            lblErrMsg.Text = string.Empty;
        }

        //表示初期化
        private void dataGridInitial(DataGridView dgv)
        {
            txtYear.BackColor = Color.Empty;
            txtMonth.BackColor = Color.Empty;
            txtNo.BackColor = Color.Empty;
            txtKubun.BackColor = Color.Empty;

            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtNo.ForeColor = Color.Navy;
            txtShozokuCode.ForeColor = Color.Navy;
            txtKubun.ForeColor = Color.Navy;

            dgv.Rows.Clear();                                      //行数をクリア
            dgv.RowCount = _MULTIGYO;                              //行数を設定
            dgv.RowsDefaultCellStyle.ForeColor = Color.Navy;       //テキストカラーの設定
            dgv.DefaultCellStyle.SelectionBackColor = Color.Empty;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Navy;
            lblNoImage.Visible = false;

            txtShinyaTl.Text = string.Empty;    // 2018/03/31
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     出勤簿データ画面表示 </summary>
        /// <param name="sIx">
        ///     データインデックス</param>
        /// <param name="rRec">
        ///     出勤簿データ配列</param>
        /// <param name="dgv">
        ///     データグリッドビューオブジェクト</param>
        ///----------------------------------------------------------------------
        private void DataShow(int sIx, string[] rRec, DataGridView dgv)
        {
            string SqlStr = string.Empty;

            lblYakushoku.Text = string.Empty;

            // 画像ファイル名
            global.pblImageFile = string.Empty;

            // データグリッドビュー初期化
            dataGridInitial(this.dataGridView1);

            //データ表示背景色初期化
            dsColorInitial(this.dataGridView1);

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            try
            {
                // 出勤簿ヘッダ
                sCom.CommandText = "select * from 出勤簿ヘッダ where ID = ?";
                sCom.Connection = dCon.cnOpen();
                sCom.Parameters.AddWithValue("@ID", rRec[sIx]);

                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    txtYear.Text = Utility.EmptytoZero(dR["年"].ToString());
                    txtMonth.Text = Utility.EmptytoZero(dR["月"].ToString());
                    txtKubun.Text = Utility.EmptytoZero(dR["勤務先区分"].ToString());

                    global.ChangeValueStatus = false;    // チェンジバリューステータス
                    txtNo.Text = string.Empty;
                    global.ChangeValueStatus = true;    // チェンジバリューステータス
                    txtNo.Text = Utility.EmptytoZero(dR["個人番号"].ToString());

                    txtShozokuCode.Text = Utility.EmptytoZero(dR["所属コード"].ToString());
                    global.pblImageFile = Properties.Settings.Default.dataPath + dR["画像名"].ToString();

                    //データ数表示
                    lblPage.Text = " (" + (cI + 1).ToString() + "/" + sID.Length.ToString() + ")";

                    if (dR["勤務先区分"].ToString() == global.KINMU_MAIN.ToString())
                    {
                        // 立替金 2016/11/16
                        txtTatekaeKin.Enabled = true;
                        txtTatekaeKin.Text = dR["立替金"].ToString();

                        // 旅費交通費 2016/11/16
                        txtRyohi.Enabled = true;
                        txtRyohi.Text = dR["旅費交通費"].ToString();

                        // その他支給 2016/11/24
                        txtSonota.Enabled = true;
                        txtSonota.Text = dR["その他支給"].ToString();
                    }
                    else
                    {
                        // 立替金 2016/11/16
                        txtTatekaeKin.Enabled = false;
                        txtTatekaeKin.Text = "";

                        // 旅費交通費 2016/11/16
                        txtRyohi.Enabled = false;
                        txtRyohi.Text = "";

                        // その他支給 2016/11/24
                        txtSonota.Enabled = false;
                        txtSonota.Text = "";
                    }
                }
                dR.Close();

                // 出勤簿明細
                sCom.CommandText = "select * from 出勤簿明細 where ヘッダID = ? order by ID";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", rRec[sIx]);
                dR = sCom.ExecuteReader();

                int r = 0;
                while (dR.Read())
                {
                    dgv[cDay, r].Value = dR["日付"];
                    dgv[cKyuka, r].Value = dR["休暇記号"];
                    dgv[cYukyu, r].Value = dR["有給記号"];
                    dgv[cSH, r].Value = dR["開始時"];
                    dgv[cSM, r].Value = dR["開始分"];
                    dgv[cEH, r].Value = dR["終了時"];
                    dgv[cEM, r].Value = dR["終了分"];
                    dgv[cKKH, r].Value = dR["規定内時"];
                    dgv[cKKM, r].Value = dR["規定内分"];
                    dgv[cKSH, r].Value = dR["深夜帯時"];
                    dgv[cKSM, r].Value = dR["深夜帯分"];
                    dgv[cTH, r].Value = dR["実働時"];
                    dgv[cTM, r].Value = dR["実働分"];

                    if (int.Parse(dR["実働編集"].ToString()) == global.flgOn)
                        dgv[cCheck, r].Value = true;
                    else dgv[cCheck, r].Value = false;

                    dgv[cID, r].Value = dR["ID"].ToString();    // 明細ＩＤ

                    r++;
                }

                dR.Close();
                sCom.Connection.Close();

                //画像表示
                // ShowImage(global.pblImageFile);

                ////////画像イメージ表示
                //////if (System.IO.File.Exists(Properties.Settings.Default.dataPath + global.pblImageFile))
                //////{
                //////    pictureBox1.Visible = true;
                //////    lblNoImage.Visible = false;

                //////    // 画像操作ボタン
                //////    btnPlus.Enabled = true;
                //////    btnMinus.Enabled = true;
                //////    button1.Enabled = true;
                //////    button2.Enabled = true;
                //////    button3.Enabled = true;
                //////    button4.Enabled = true;

                //////    // 画像を表示する
                //////    ImageGraphicsPaint(pictureBox1, Properties.Settings.Default.dataPath + global.pblImageFile,
                //////                       global.ZOOM_RATE, global.ZOOM_RATE, 0, 0);
                //////    pictureBox1.Refresh();
                //////}
                //////else
                //////{
                //////    pictureBox1.Visible = false;
                //////    lblNoImage.Visible = true;

                //////    // 画像操作ボタン
                //////    btnPlus.Enabled = false;
                //////    btnMinus.Enabled = false;
                //////    button1.Enabled = false;
                //////    button2.Enabled = false;
                //////    button3.Enabled = false;
                //////    button4.Enabled = false;
                //////}

                // ヘッダ情報
                txtYear.ReadOnly = false;
                txtMonth.ReadOnly = false;
                txtShozokuCode.ReadOnly = false;
                txtNo.ReadOnly = false;

                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum = rRec.Length - 1;
                hScrollBar1.Value = sIx;
                hScrollBar1.LargeChange = 1;
                hScrollBar1.SmallChange = 1;

                //移動ボタン制御
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;

                //最初のレコード
                if (sIx == 0)
                {
                    btnBefore.Enabled = false;
                    btnFirst.Enabled = false;
                }

                //最終レコード
                if ((sIx + 1) == rRec.Length)
                {
                    btnNext.Enabled = false;
                    btnEnd.Enabled = false;
                }

                //カレントセル選択状態としない
                dgv.CurrentCell = null;
                //dataGridView2.CurrentCell = null;

                // その他のボタンを有効とする
                button5.Enabled = true;
                btnErrCheck.Visible = true;
                btnDataMake.Visible = true;
                btnDel.Visible = true;

                // データグリッドビュー編集
                dataGridView1.ReadOnly = false;

                // 編集チェックオン・オフ
                linkLabel1.Visible = true;
                linkLabel2.Visible = true;

                //エラー情報表示
                ErrShow();

                // 深夜勤務の誤表示対策のため再度計算させて表示 2018/03/31
                txtShinyaTl.Text = getShinyaTime().ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void LastDataShow(string sID, DataGridView dgv)
        {
            string SqlStr = string.Empty;

            // 画像ファイル名
            global.pblImageFile = string.Empty;

            // データグリッドビュー初期化
            dataGridInitial(this.dataGridView1);

            //データ表示背景色初期化
            dsColorInitial(this.dataGridView1);

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR = null;

            try
            {
                // 過去出勤簿ヘッダ
                sCom.CommandText = "select * from 過去出勤簿ヘッダ where ID = ?";
                sCom.Parameters.AddWithValue("@ID", sID);

                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    //_PCADBName = dR["画像名"].ToString().Substring(0, 17);
                    txtKubun.Text = Utility.EmptytoZero(dR["勤務先区分"].ToString());    // 2016/11/23
                    txtYear.Text = (int.Parse(dR["年"].ToString()) - Utility.GetRekiHosei()).ToString();
                    txtMonth.Text = Utility.EmptytoZero(dR["月"].ToString());
                    txtNo.Text = Utility.EmptytoZero(dR["個人番号"].ToString());
                    txtShozokuCode.Text = Utility.EmptytoZero(dR["所属コード"].ToString());
                    lblShozoku.Text = Utility.NulltoStr(dR["所属名"].ToString());

                    lblName.Text = dR["氏名"].ToString();

                    if (dR["給与区分"].ToString() == global.STATUS_PART.ToString())
                    {
                        lblYakushoku.Text = "パート";
                    }
                    else
                    {
                        lblYakushoku.Text = "社員";
                    }

                    global.pblImageFile = Properties.Settings.Default.tifPath + dR["画像名"].ToString();

                    //データ数表示
                    lblPage.Text = " (" + (cI + 1).ToString() + "/" + sID.Length.ToString() + ")";

                    // 集計項目表示
                    txtShukkinTl.Text = dR["出勤日数合計"].ToString();    // 出勤日数
                    txtYukyuHiTl.Text = dR["有休日数合計"].ToString();    // 有休日数
                    txtYukyuTmTl.Text = dR["有休時間合計"].ToString();    // 有給時間
                    txtTokkyuTl.Text = dR["特休日数合計"].ToString();     // 特休日数
                    txtKekkinTl.Text = dR["欠勤日数合計"].ToString();     // 欠勤日数
                    txtChisouTl.Text = dR["遅刻早退回数"].ToString();     // 遅刻早退回数
                    txtFurikyuTl.Text = dR["振休日数合計"].ToString();    // 振休日数合計
                    txtFurideTl.Text = dR["振出日数合計"].ToString();     // 振出日数合計
                    txtRhTl.Text = dR["総労働"].ToString();               // 総労働時間・時
                    txtRmTl.Text = dR["総労働分"].ToString();             // 総労働時間･分
                    txtZanHTl.Text = dR["残業時"].ToString();             // 残業時間・時
                    txtZanMTl.Text = dR["残業分"].ToString();             // 残業時間･分
                    txtShinyaTl.Text = dR["深夜勤務時間合計"].ToString();  // 深夜勤務時間合計
                    txtKiteiTl.Text = dR["月間規定勤務時間"].ToString();    // 月間規定勤務時間
                    txtPtSouwaku.Text = dR["パート労働時間総枠"].ToString();    // パート労働時間総枠                    
                    txtTatekaeKin.Text = dR["立替金"].ToString();  // 立替金 2016/11/23                    
                    txtRyohi.Text = dR["旅費交通費"].ToString(); // 旅費交通費 2016/11/23
                    txtSonota.Text = dR["その他支給"].ToString(); // その他支給 2016 /11/24
                }
                dR.Close();

                // 過去出勤簿明細
                sCom.CommandText = "select * from 過去出勤簿明細 where ヘッダID = ? order by ID";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                dR = sCom.ExecuteReader();

                int r = 0;
                while (dR.Read())
                {
                    dgv[cDay, r].Value = dR["日付"];
                    dgv[cKyuka, r].Value = dR["休暇記号"];
                    dgv[cYukyu, r].Value = dR["有給記号"];
                    dgv[cSH, r].Value = dR["開始時"];
                    dgv[cSM, r].Value = dR["開始分"];
                    dgv[cEH, r].Value = dR["終了時"];
                    dgv[cEM, r].Value = dR["終了分"];
                    dgv[cKKH, r].Value = dR["規定内時"];
                    dgv[cKKM, r].Value = dR["規定内分"];
                    dgv[cKSH, r].Value = dR["深夜帯時"];
                    dgv[cKSM, r].Value = dR["深夜帯分"];
                    dgv[cTH, r].Value = dR["実働時"];
                    dgv[cTM, r].Value = dR["実働分"];

                    if (int.Parse(dR["実働編集"].ToString()) == global.flgOn)
                        dgv[cCheck, r].Value = true;
                    else dgv[cCheck, r].Value = false;

                    dgv[cID, r].Value = dR["ID"].ToString();    // 明細ＩＤ

                    r++;
                }

                dR.Close();
                sCom.Connection.Close();

                //画像表示
                ShowImage(global.pblImageFile);

                // ヘッダ情報
                txtKubun.ReadOnly = true;
                txtYear.ReadOnly = true;
                txtMonth.ReadOnly = true;
                txtShozokuCode.ReadOnly = true;
                txtNo.ReadOnly = true;
                txtTatekaeKin.ReadOnly = true;
                txtRyohi.ReadOnly = true;
                txtSonota.ReadOnly = true;

                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum = 0;
                hScrollBar1.Value = 0;
                hScrollBar1.LargeChange = 1;
                hScrollBar1.SmallChange = 1;

                //移動ボタン制御
                btnFirst.Enabled = false;
                btnNext.Enabled = false;
                btnBefore.Enabled = false;
                btnEnd.Enabled = false;
                
                //カレントセル選択状態としない
                dgv.CurrentCell = null;
                //dataGridView2.CurrentCell = null;

                // その他のボタンを有効とする
                button5.Enabled = false;
                btnErrCheck.Visible = false;
                btnDataMake.Visible = false;
                btnDel.Visible = false;

                // データグリッドビュー編集
                dataGridView1.ReadOnly = true;

                // 編集チェックオン・オフ
                linkLabel1.Visible = false;
                linkLabel2.Visible = false;

                //エラー情報表示
                //ErrShow();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     画像を表示する</summary>
        /// <param name="pic">
        ///     pictureBoxオブジェクト</param>
        /// <param name="imgName">
        ///     イメージファイルパス</param>
        /// <param name="fX">
        ///     X方向のスケールファクター</param>
        /// <param name="fY">
        ///     Y方向のスケールファクター</param>
        ///----------------------------------------------------------------------
        private void ImageGraphicsPaint(PictureBox pic, string imgName, float fX, float fY, int RectDest, int RectSrc)
        {
            Image _img = Image.FromFile(imgName);
            Graphics g = Graphics.FromImage(pic.Image);
            
            // 各変換設定値のリセット
            g.ResetTransform();

            // X軸とY軸の拡大率の設定
            g.ScaleTransform(fX, fY); 
            
            // 画像を表示する
            g.DrawImage(_img, RectDest, RectSrc);

            // 現在の倍率,座標を保持する
            global.ZOOM_NOW = fX;
            global.RECTD_NOW = RectDest;
            global.RECTS_NOW = RectSrc;
        }

        ///データ表示エリア背景色初期化
        private void dsColorInitial(DataGridView dgv)
        {
            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtNo.BackColor = Color.White;
            txtShozokuCode.BackColor = Color.White;

            for (int i = 0; i < _MULTIGYO; i++)
            {
                dgv.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
            }
        }

        private void txtNo_TextChanged(object sender, EventArgs e)
        {
            // 過去データ表示のときは何もしない
            if (dID != string.Empty) return;

            // チェンジバリューステータス
            if (!global.ChangeValueStatus) return; 

            // 表示欄初期化
            this.txtShozokuCode.Text = string.Empty;
            this.lblShozoku.Text = string.Empty;
            this.lblName.Text = string.Empty;

            string zCode = string.Empty;
            string zName = string.Empty;

            this.lblName.Text = Utility.ComboShain.getXlsSname(xS, Utility.StrtoInt(txtNo.Text), out zCode, out zName);
            //txtShozokuCode.Text = zCode;    // 勤務先コード
            //lblShozoku.Text = zName;        // 勤務先名

            // 社員・パート区分表示
            if (lblName.Text != string.Empty)
            {
                // 社員・パート
                if (_YakushokuType == global.STATUS_SHAIN)
                {
                    lblYakushoku.Text = "社員";
                }
                else if (_YakushokuType == global.STATUS_PART)
                {
                    lblYakushoku.Text = "パート";
                }
                else
                {
                    lblYakushoku.Text = string.Empty;
                }
            }

            // 月間勤務時間取得 2016/11/16
            getMonthWorkTime();

            // パート変形労働時間制における労働時間総枠を取得する
            if (txtKiteiTl.Text == string.Empty)
            {
                GetPartSouwakuLoad();
            }

            // 休日再表示
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                YoubiSet(i);
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     月間勤務時間取得 2016/11/16 </summary> 
        ///----------------------------------------------------------
        private void getMonthWorkTime()
        {
            StringBuilder sb = new StringBuilder();

            // 月間勤務時間取得
            SysControl.SetDBConnect mdb = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = mdb.cnOpen();

            OleDbDataReader dR;

            sb.Clear();
            sb.Append("select 月間勤務時間 from 月給者勤務時間 ");
            sb.Append("where データ領域 = ? and 社員番号 = ?");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.AddWithValue("@dID", Utility.StrtoInt(_grpID));
            sCom.Parameters.AddWithValue("@ID", Utility.StrtoInt(txtNo.Text));

            dR = sCom.ExecuteReader();

            txtKiteiTl.Enabled = true;
            txtKiteiTl.Text = string.Empty;
            txtPtSouwaku.Enabled = false;
            txtPtSouwaku.Text = string.Empty;

            while (dR.Read())
            {
                txtKiteiTl.Text = dR["月間勤務時間"].ToString();
            }

            dR.Close();
            sCom.Connection.Close();
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     パート変形労働時間制における労働時間総枠を取得する </summary>
        /// ----------------------------------------------------------------------
        private void GetPartSouwakuLoad()
        {
            if (_YakushokuType == global.STATUS_PART)
            {
                string sDate = txtYear.Text + "/" + txtMonth.Text + "/01";
                DateTime eDate;
                if (DateTime.TryParse(sDate, out eDate))
                {
                    double sw = getPartSouwaku(Utility.StrtoInt(Utility.EmptytoZero(txtYear.Text)) + Utility.GetRekiHosei(),
                                                  Utility.StrtoInt(txtMonth.Text));

                    txtPtSouwaku.Enabled = true;
                    txtPtSouwaku.Text = sw.ToString();
                    txtKiteiTl.Enabled = false;
                    txtKiteiTl.Text = string.Empty;
                }
                else
                {
                    txtPtSouwaku.Enabled = false;
                    txtPtSouwaku.Text = string.Empty;
                }
            }
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     パート変形労働時間制における労働時間総枠を取得する </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <returns>
        ///     パート変形労働時間制における労働時間総枠</returns>
        /// ----------------------------------------------------------------------
        private double GetPartSouwakuLoad(string yy, string mm)
        {
            double rtn = 0;

            if (_YakushokuType == 1)
            {
                string sDate = yy + "/" + mm + "/01";
                DateTime eDate;
                if (DateTime.TryParse(sDate, out eDate))
                {
                    double sw = getPartSouwaku(Utility.StrtoInt(Utility.EmptytoZero(yy)) + Properties.Settings.Default.RekiHosei,
                                                  Utility.StrtoInt(mm));

                    rtn = sw;
                }
                else
                {
                    rtn = 0;
                }
            }

            return rtn;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
        private void txtKubun_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '1' && e.KeyChar != '2' && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                if (dataGridView1.CurrentCell.ColumnIndex != 19)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        /// <summary>
        /// 曜日をセットする
        /// </summary>
        /// <param name="tempRow">MultiRowのindex</param>
        private void YoubiSet(int tempRow)
        {
            string sDate;
            DateTime eDate;
            Boolean bYear = false;
            Boolean bMonth = false;

            //年月を確認
            if (txtYear.Text != string.Empty)
            {
                if (Utility.NumericCheck(txtYear.Text))
                {
                    if (int.Parse(txtYear.Text) > 0)
                    {
                        bYear = true;
                    }
                }
            }

            if (txtMonth.Text != string.Empty)
            {
                if (Utility.NumericCheck(txtMonth.Text))
                {
                    if (int.Parse(txtMonth.Text) >= 1 && int.Parse(txtMonth.Text) <= 12)
                    {
                        for (int i = 0; i < _MULTIGYO; i++)
                        {
                            bMonth = true;
                        }
                    }
                }
            }

            //年月の値がfalseのときは曜日セットは行わずに終了する
            if (bYear == false || bMonth == false) return;

            //行の色を初期化
            dataGridView1.Rows[tempRow].DefaultCellStyle.BackColor = Color.Empty;

            //Nullか？
            dataGridView1[cWeek, tempRow].Value = string.Empty;
            if (dataGridView1[cDay, tempRow].Value != null) 
            {
                if (dataGridView1[cDay, tempRow].Value.ToString() != string.Empty)
                {
                    if (Utility.NumericCheck(dataGridView1[cDay, tempRow].Value.ToString()))
                    {
                        {
                            sDate = (int.Parse(Utility.EmptytoZero(txtYear.Text)) + Utility.GetRekiHosei()).ToString() + "/" +
                                               Utility.EmptytoZero(txtMonth.Text) + "/" +
                                               Utility.EmptytoZero(dataGridView1[cDay, tempRow].Value.ToString());
                            
                            // 存在する日付と認識された場合、曜日を表示する
                            if (DateTime.TryParse(sDate, out eDate))
                            {
                                dataGridView1[cWeek, tempRow].Value = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);

                                // 休日背景色設定・日曜日
                                if (dataGridView1[cWeek, tempRow].Value.ToString() == "日")
                                    dataGridView1.Rows[tempRow].DefaultCellStyle.BackColor = Color.MistyRose;
                                else
                                {
                                    //祝祭日なら曜日の背景色を変える
                                    for (int j = 0; j < Holiday.Length; j++)
                                    {
                                        // 休日登録されている
                                        if (Holiday[j].hDate == eDate)
                                        {
                                            // 社員またはパートが各々休日対象となっている
                                            if (_YakushokuType != 1 && Holiday[j].Gekkyuu == true ||
                                                _YakushokuType == 1 && Holiday[j].Jikyuu == true)
                                            {
                                                dataGridView1.Rows[tempRow].DefaultCellStyle.BackColor = Color.MistyRose;
                                                break;
                                            }
                                            else dataGridView1.Rows[tempRow].DefaultCellStyle.BackColor = Color.White;
                                        }
                                    }
                                }

                                // 時刻区切り文字
                                dataGridView1[cSE, tempRow].Value = ":";
                                dataGridView1[cEE, tempRow].Value = ":";
                                dataGridView1[cKKE, tempRow].Value = ":";
                                dataGridView1[cKSE, tempRow].Value = ":";
                                dataGridView1[cTE, tempRow].Value = ":";
                            }
                        }
                    }
                }
             }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!global.ChangeValueStatus) return;
            if (e.RowIndex < 0) return;

            string colName = dataGridView1.Columns[e.ColumnIndex].Name;

            if (colName == cDay) YoubiSet(e.RowIndex);  // 日付

            // 過去データ表示のときは終了
            if (dID != string.Empty) return;

            // 出勤日数
            txtShukkinTl.Text = getWorkDays(_YakushokuType);

            // 休暇記号
            if (colName == cKyuka) 
            {
                txtTokkyuTl.Text = getKyukaTotal(global.TOKUBETSU_KYUKA);
                txtKekkinTl.Text = getKyukaTotal(global.KEKKIN_KYUKA);
                txtChisouTl.Text = getKyukaTotal(global.CHISOU_KYUKA);
                txtFurikyuTl.Text = getKyukaTotal(global.FURIKYU_KYUKA);
                txtFurideTl.Text = getKyukaTotal(global.FURIDE_KYUKA);
            }

            // 有給記号
            if (colName == cYukyu) 
            {
                txtYukyuHiTl.Text = getYukyuTotal(0);
                //txtYukyuTmTl.Text = getYukyuTotal(1);
            }

            // 深夜勤務
            if (colName == cSH || colName == cSM || colName == cEH || colName == cEM || 
                colName == cKSH || colName == cKSM)
            {
                txtShinyaTl.Text = getShinyaTime().ToString();
            }

            // 実労働時間
            if (colName == cTH || colName == cTM || colName == cYukyu)
            {
                double w = 0;

                // パートタイマーは月間実労働時間合計を計算します
                if (_YakushokuType == global.STATUS_PART)
                {
                    w = getWorkTime();
                    txtRhTl.Text = System.Math.Floor(w / 60).ToString();
                    txtRmTl.Text = (w % 60).ToString();
                }
                else
                {
                    txtRhTl.Text = string.Empty;
                    txtRmTl.Text = string.Empty;
                }

                // 残業時間・社員、一部パートの月間規定勤務時間設定者
                if (Utility.NulltoStr(txtKiteiTl.Text) != string.Empty)
                {
                    w = getZangyoTime();
                    txtZanHTl.Text = System.Math.Floor(w / 60).ToString();
                    txtZanMTl.Text = (w % 60).ToString();
                }
                else if (_YakushokuType == global.STATUS_PART)   // パートタイマー
                {
                    w = getZangyoPart();
                    txtZanHTl.Text = System.Math.Floor(w / 60).ToString();
                    txtZanMTl.Text = (w % 60).ToString();
                }
            }

            // 実労働時間編集チェック
            if (colName == cCheck)
            {
                if (dataGridView1[cCheck, e.RowIndex].Value.ToString() == "True")
                {
                    dataGridView1[cTH, e.RowIndex].ReadOnly = false;
                    dataGridView1[cTM, e.RowIndex].ReadOnly = false;
                }
                else
                {
                    dataGridView1[cTH, e.RowIndex].ReadOnly = true;
                    dataGridView1[cTM, e.RowIndex].ReadOnly = true;
                }
            }
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     与えられた休暇記号に該当する休暇日数取得 </summary>
        /// <param name="kigou">
        ///     休暇記号</param>
        /// <returns>
        ///     休暇日数</returns>
        ///---------------------------------------------------------
        private string getKyukaTotal(string kigou)
        {
            int days = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[cKyuka, i].Value != null)
                {
                    if (dataGridView1[cKyuka, i].Value.ToString() == kigou)
                        days++;
                }
            }

            return days.ToString();
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     有給休暇日数・時間取得 </summary>
        /// <returns>
        ///     有給日数</returns>
        /// <param name="Status">
        ///     0:日数、1:時間</param>
        ///--------------------------------------------------------------
        private string getYukyuTotal(int Status)
        {
            double days = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[cYukyu, i].Value != null)
                {
                    if (Status == 0)    // 有給日数 2016/11/17
                    {
                        if (dataGridView1[cYukyu, i].Value.ToString() == global.ZENNICHI_YUKYU)
                        {
                            days++;
                        }
                        else if (dataGridView1[cYukyu, i].Value.ToString() == global.HANNICHI_YUKYU)
                        {
                            days += (double)(0.5);
                        }
                    }
                    else if (Status == 1)   // 有給時間
                    {
                        if (dataGridView1[cYukyu, i].Value.ToString() != global.ZENNICHI_YUKYU &&
                            dataGridView1[cYukyu, i].Value.ToString() != string.Empty)
                            days += int.Parse(dataGridView1[cYukyu, i].Value.ToString());
                    }
                }
            }

            return days.ToString();
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     総労働時間取得 </summary>
        /// <returns>
        ///     総労働時間・分</returns>
        ///----------------------------------------------------------
        private int getWorkTime()
        {
            int wHour = 0;
            int wMin = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                wHour += Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cTH, i].Value));
                wMin += Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cTM, i].Value));
            }

            return (wHour * 60 + wMin);
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     全日有給休暇を含まない総労働時間取得：2017/06/22 
        ///     半日有給は実労時間の1/2を労働時間とする：2018/04/05</summary>
        /// <returns>
        ///     総労働時間・分</returns>
        ///-----------------------------------------------------------------------------
        private double getWorkTimeNotYukyu()
        {
            double wHour = 0;
            double wMin = 0;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                // 全日有休は対象外
                if (Utility.NulltoStr(dataGridView1[cYukyu, i].Value) == global.ZENNICHI_YUKYU)
                {
                    continue;
                }

                // 半日有給は実労時間の1/2を当日の労働時間とする：2018/04/05
                if (Utility.NulltoStr(dataGridView1[cYukyu, i].Value) == global.HANNICHI_YUKYU)
                {
                    double hh = Utility.StrtoDouble(Utility.NulltoStr(dataGridView1[cTH, i].Value)) / 2;
                    double mm = Utility.StrtoDouble(Utility.NulltoStr(dataGridView1[cTM, i].Value)) / 2;

                    //MessageBox.Show((i + 1) + "日 半休 " + hh + ":" + mm);

                    wHour += hh;
                    wMin += mm;
                }
                else
                {
                    wHour += Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cTH, i].Value));
                    wMin += Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cTM, i].Value));
                }
            }

            return (wHour * 60 + wMin);
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     残業時間取得 </summary>
        /// <returns>
        ///     残業時間・分</returns>
        ///-----------------------------------------------------------------------
        private int getZangyoTime()
        {
            int wHour = 0;
            int wMin = 0;
            DateTime dt;
            double spanMin = 0;
            int zan = 0;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                // 全日有休は対象外：2017/06/22
                if (Utility.NulltoStr(dataGridView1[cYukyu, i].Value) == global.ZENNICHI_YUKYU)
                {
                    continue;
                }

                wHour = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cTH, i].Value));
                wMin = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cTM, i].Value));

                if (wHour > Properties.Settings.Default.DayWorkTime || 
                    (wHour == Properties.Settings.Default.DayWorkTime && wMin > 0))
                {
                    if (DateTime.TryParse(wHour.ToString() + ":" + wMin.ToString(), out dt))
                    {
                        spanMin += Utility.GetTimeSpan(global.dt0800, dt).TotalMinutes;
                    }
                }
            }

            if (spanMin > Utility.StrtoInt(Utility.NulltoStr(txtKiteiTl.Text)) * 60)
                zan = (int)(spanMin -  Utility.StrtoInt(Utility.NulltoStr(txtKiteiTl.Text)) * 60);

            return zan;
        }

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     残業時間取得 月間規定勤務時間有対象・複数枚勤務票対応</summary>
        /// <param name="wsID">
        ///     社員ID</param>
        /// <param name="kiteiTM">
        ///     月間規定勤務時間</param>
        /// <returns>
        ///     残業時間・分</returns>
        ///     2014/10/28
        ///---------------------------------------------------------------------------
        private int getZangyoTimeTotal(string wsID, int kiteiTM)
        {
            int wHour = 0;
            int wMin = 0;
            DateTime dt;
            double spanMin = 0;
            int zan = 0;

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;
            sCom.Connection = dCon.cnOpen();

            // 出勤簿明細
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT 出勤簿明細.* from 出勤簿ヘッダ inner join 出勤簿明細 ");
            sb.Append("on 出勤簿明細.ヘッダID = 出勤簿ヘッダ.ID ");
            sb.Append("where 出勤簿ヘッダ.社員ID = ? order by 出勤簿明細.ID");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@ID", wsID);
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                wHour = Utility.StrtoInt(Utility.NulltoStr(dR["実働時"]));
                wMin = Utility.StrtoInt(Utility.NulltoStr(dR["実働分"]));

                if (wHour > Properties.Settings.Default.DayWorkTime ||
                    (wHour == Properties.Settings.Default.DayWorkTime && wMin > 0))
                {
                    if (DateTime.TryParse(wHour.ToString() + ":" + wMin.ToString(), out dt))
                    {
                        spanMin += Utility.GetTimeSpan(global.dt0800, dt).TotalMinutes;
                    }
                }
            }

            dR.Close();
            sCom.Connection.Close();

            if (spanMin > kiteiTM * 60)
                zan = (int)(spanMin - kiteiTM * 60);

            return zan;
        }

        /// -----------------------------------------------------------
        /// <summary>
        ///     残業時間計算・パート </summary>
        /// <returns>
        ///     残業時間（分）</returns>
        /// -----------------------------------------------------------
        private int getZangyoPart()
        {
            int zan = 0;

            //// 総労働時間・分
            //double wMin = Utility.StrtoInt(Utility.NulltoStr(txtRhTl.Text)) * 60 +
            //       Utility.StrtoInt(Utility.NulltoStr(txtRmTl.Text));

            //double w = Utility.ToRoundDown(wMin / 60, 1);

            // 全日有給休暇を含まない総労働時間を取得：2017/06/21
            // 半日有給休暇は総労働時間の1/2を取得：2018/04/05
            //double w = Utility.ToRoundDown(getWorkTimeNotYukyu() / 60, 1);  // 2019/05/31 端数差異発生のため、コメント化
            double w = getWorkTimeNotYukyu();

            // 当月の労働時間総枠を取得する
            //double sw = getPartSouwaku(Utility.StrtoInt(Utility.EmptytoZero(txtYear.Text)) + Properties.Settings.Default.RekiHosei,
            //Utility.StrtoInt(txtMonth.Text));
            GetPartSouwakuLoad();

            // パート変形労働時間制における労働時間総枠を超えた分を残業とする  // 2019/05/31 コメント化
            // double sw = Utility.StrtoDouble(Utility.NulltoStr(txtPtSouwaku.Text)); 
            //if (w > sw)
            //{
            //    zan = (int)System.Math.Floor((w - sw) * 60);
            //}

            // パート変形労働時間制における労働時間総枠を超えた分を残業とする  // 2019/05/31 
            double sw = Utility.StrtoDouble(Utility.NulltoStr(txtPtSouwaku.Text)) * 60; // 分単位に変換 2019/05/31
            if (w > sw)
            {
                //zan = (int)System.Math.Floor((w - sw) * 60);  // 2019/05/31 コメント化
                zan = (int)(w - sw);    // 分単位で超過時間を計算
            }

            //MessageBox.Show(w.ToString() + "  :  " + zan.ToString());

            return zan;
        }

        /// ----------------------------------------------------------------------------------
        /// <summary>
        ///     残業時間計算・パート 同一社員・複数勤務票対応　2014/10/24 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="h">
        ///     総労働時間・時</param>
        /// <param name="m">
        ///     総労働時間・分</param>
        /// <returns>
        ///     残業時間（分）</returns>
        /// ----------------------------------------------------------------------------------
        private double getZangyoPartTotal(string yy, string mm, string h, string m)
        {
            double zan = 0;

            // 総労働時間・分
            double wMin = Utility.StrtoInt(Utility.NulltoStr(h)) * 60 + 
                   Utility.StrtoInt(Utility.NulltoStr(m));

            double w = Utility.ToRoundDown(wMin / 60, 1);

            // 当月の労働時間総枠を取得する
            double sw = GetPartSouwakuLoad(yy, mm);

            // パート変形労働時間制における労働時間総枠を超えた分を残業とする 
            //double sw = Utility.StrtoDouble(Utility.NulltoStr(txtPtSouwaku.Text));
            if (w > sw)
            {
                zan = System.Math.Floor((w - sw) * 60);
            }

            return zan;
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     パート変形労働時間制における労働時間総枠を取得する </summary>
        /// <param name="y">
        ///     該当年</param>
        /// <param name="m">
        ///     該当月</param>
        /// <returns>
        ///     総枠労働時間数</returns>
        /// -------------------------------------------------------------------
        private double getPartSouwaku(int y, int m)
        {
            // 当月の暦日数
            int rn = DateTime.DaysInMonth(y, m);

            // パート変形労働時間制における労働時間総枠を返す
            for (int i = 0; i < 4; i++)
            {
                if (ptLimitTm[i, 0] == rn)
                {
                    return ptLimitTm[i, 1];
                }
            }

            return 0;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     出勤日数取得 </summary>
        /// <returns>
        ///     出勤日数</returns>
        ///----------------------------------------------------------
        private string getWorkDays(int yaku)
        {
            // 出勤日数
            int sDays = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                // 勤務時間が記入されている行
                if ((dataGridView1[cSH, i].Value != null && dataGridView1[cSH, i].Value.ToString() != string.Empty) && 
                    (dataGridView1[cSM, i].Value != null && dataGridView1[cSM, i].Value.ToString() != string.Empty))
                {
                    // 開始時間が24時台以外のもの（24時台は前日からの通し勤務とみなし出勤日数に加えない）
                    if (dataGridView1[cSH, i].Value.ToString() != "24")
                    {
                        // 社員
                        if (yaku == global.STATUS_SHAIN)
                        {
                            sDays++;
                        }
                        else if (dataGridView1[cYukyu, i].Value != null)    // パート：終日有休以外のときは出勤日数としてカウントする
                        {
                            if (dataGridView1[cYukyu, i].Value.ToString() != global.ZENNICHI_YUKYU)
                            {
                                sDays++;
                            }
                        }
                    }
                }
            }

            return sDays.ToString();
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間取得(22:00～05:00)　2014/03/04に算出方法を改訂 </summary>
        /// <returns>
        ///     深夜勤務時間・分</returns>
        ///--------------------------------------------------------------------------
        private double getShinyaTime()
        {
            //int wHour = 0;
            //int wMin = 0;
            int wHourk = 0;
            int wMink = 0;
            int sKyukei = 0;

            int sHour = 0;
            int sMin = 0;
            int eHour = 0;
            int eMin = 0;

            //DateTime stTM;
            //DateTime edTM;
            DateTime cTM;

            DateTime pSTM;  // 記入された開始時間
            DateTime pETM;  // 記入された終了時間

            double spanMin = 0;
            
            // 深夜時間帯 : 2018/03/31
            DateTime dt2200 = DateTime.Parse("22:00");
            DateTime dt0500 = DateTime.Parse("05:00");

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                // 2014/03/04
                if (Utility.NulltoStr(dataGridView1[cSH, i].Value) != string.Empty &&
                    Utility.NulltoStr(dataGridView1[cSM, i].Value) != string.Empty &&
                    Utility.NulltoStr(dataGridView1[cEH, i].Value) != string.Empty &&
                    Utility.NulltoStr(dataGridView1[cEM, i].Value) != string.Empty)
                {
                    // 開始時刻を取得　2014/03/04
                    sHour = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cSH, i].Value));
                    sMin = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cSM, i].Value));
                    if (sHour == 24) sHour = 0;
                    if (DateTime.TryParse(sHour.ToString() + ":" + sMin.ToString(), out cTM))
                    {
                        pSTM = cTM; // 開始時刻を取得
                    }
                    else continue;  // 時刻型に変換できなかった場合はネグる
                    
                    // 終了時刻を取得　2014/03/04
                    eHour = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cEH, i].Value));
                    eMin = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cEM, i].Value));

                    // 24:00記入は23:59に変更する
                    if (eHour == 24 && eMin == 0)
                    {
                        eHour = 23;
                        eMin = 59;
                    }

                    if (DateTime.TryParse(eHour.ToString() + ":" + eMin.ToString(), out cTM))
                    {
                        pETM = cTM; // 終了時刻を取得
                    }
                    else continue;  // 時刻型に変換できなかった場合はネグる

                    // 開始が５：００以前のとき
                    if (pSTM < dt0500)
                    {
                        // 終了時刻が午前5時以降か
                        if (pETM >= dt0500)
                            spanMin += Utility.GetTimeSpan(pSTM, dt0500).TotalMinutes;
                        else spanMin += Utility.GetTimeSpan(pSTM, pETM).TotalMinutes;
                    }
                    
                    // 終了が２２：００以降のとき
                    if (pETM > dt2200)
                    {
                        // 開始時刻が22時以前か
                        if (pSTM < dt2200)
                            spanMin += Utility.GetTimeSpan(dt2200, pETM).TotalMinutes;
                        else spanMin += Utility.GetTimeSpan(pSTM, pETM).TotalMinutes;
                    }

                    // 終了が24:00のときは23:59まで計算なので1分加算する
                    if (Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cEH, i].Value)) == 24 &&
                        Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cEM, i].Value)) == 0)
                    {
                        spanMin += 1;
                    }
                    

                    //// 開始が５：００以前のとき
                    //if (Utility.NulltoStr(dataGridView1[cSH, i].Value) != string.Empty &&
                    //    Utility.NulltoStr(dataGridView1[cSM, i].Value) != string.Empty)
                    //{
                    //    wHour = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cSH, i].Value));
                    //    wMin = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cSM, i].Value));

                    //    if (wHour == 24) wHour = 0;

                    //    if (wHour < 5 && wMin < 60)
                    //    {
                    //        // 深夜勤務時間
                    //        stTM = DateTime.Parse(wHour.ToString() + ":" + wMin.ToString());
                    //        spanMin += Utility.GetTimeSpan(stTM, global.dt0500).TotalMinutes;
                    //    }
                    //}

                    //// 終了が２２：００以降のとき
                    //if (Utility.NulltoStr(dataGridView1[cEH, i].Value) != string.Empty &&
                    //    Utility.NulltoStr(dataGridView1[cEM, i].Value) != string.Empty)
                    //{
                    //    wHour = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cEH, i].Value));
                    //    wMin = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cEM, i].Value));

                    //    if (wHour >= 22)
                    //    {
                    //        // 深夜勤務時間
                    //        //sHour = (wHour - 22) * 60 + wMin;

                    //        if (wHour < 25 && wMin < 60)
                    //        {
                    //            if (wHour < 24)
                    //            {
                    //                edTM = DateTime.Parse(wHour.ToString() + ":" + wMin.ToString());

                    //                // 開始時間が22時以降に対応 2014/03/04
                    //                //if (stTM > global.dt2200)
                    //                spanMin += Utility.GetTimeSpan(global.dt2200, edTM).TotalMinutes;
                    //            }
                    //            // 24:00のときは23:59まで計算して1分加算する
                    //            else if (wMin == 0)
                    //            {
                    //                edTM = DateTime.Parse("23:59");
                    //                spanMin += Utility.GetTimeSpan(global.dt2200, edTM).TotalMinutes + 1;
                    //            }
                    //        }
                    //    }
                    //}

                    // 深夜帯休憩時間
                    wHourk = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cKSH, i].Value));
                    wMink = Utility.StrtoInt(Utility.NulltoStr(dataGridView1[cKSM, i].Value));
                    sKyukei = wHourk * 60 + wMink;

                    // 深夜帯休憩時間を除して深夜勤務時間を求める 2014/03/04
                    if (spanMin >= sKyukei)
                    {
                        spanMin -= sKyukei;
                    }
                }
            }

            return spanMin;
        }

        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty) btnRtn.Focus();
        }

        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (cI + 1 < sID.Length)
            {
                cI++;
                DataShow(cI, sID, dataGridView1);
            }   
        }


        ///---------------------------------------------------------------------
        /// <summary>
        ///     カレントデータの更新 </summary>
        /// <param name="iX">
        ///     カレントレコードのインデックス</param>
        ///---------------------------------------------------------------------
        private void CurDataUpDate(int iX)
        {
            //カレントデータを更新する
            string mySql = string.Empty;

            // エラーメッセージ
            string errMsg = string.Empty;

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            // 出勤簿ヘッダテーブル
            mySql += "update 出勤簿ヘッダ set ";
            mySql += "社員ID=?,個人番号=?,氏名=?,年=?,月=?,所属コード=?,所属名=?,給与区分=?,";
            mySql += "出勤日数合計=?,有休日数合計=?,有休時間合計=?,特休日数合計=?,振休日数合計=?,";
            mySql += "振出日数合計=?,遅刻早退回数=?,欠勤日数合計=?,";
            mySql += "実稼動日数合計=?,総労働=?,総労働分=?,残業時=?,残業分=?,深夜勤務時間合計=?,";
            mySql += "月間規定勤務時間=?,パート労働時間総枠=?,確認=?,更新年月日=?,立替金=?,旅費交通費=?,勤務先区分=?,その他支給=?";
            mySql += "where ID = ?";

            errMsg = "出勤簿ヘッダテーブル更新";

            sCom.CommandText = mySql;
            sCom.Parameters.AddWithValue("@ShainId", _ShainID);
            sCom.Parameters.AddWithValue("@Code", Utility.NulltoStr(txtNo.Text).PadLeft(5, '0'));
            sCom.Parameters.AddWithValue("@Name", Utility.NulltoStr(lblName.Text));
            sCom.Parameters.AddWithValue("@Year", Utility.StrtoInt(Utility.NulltoStr(txtYear.Text)));
            sCom.Parameters.AddWithValue("@Month", Utility.StrtoInt(Utility.NulltoStr(txtMonth.Text)));
            sCom.Parameters.AddWithValue("@Shozoku", Utility.NulltoStr(txtShozokuCode.Text));
            sCom.Parameters.AddWithValue("@SzName", Utility.NulltoStr(lblShozoku.Text).Trim());
            sCom.Parameters.AddWithValue("@Yakushoku", _YakushokuType);
            sCom.Parameters.AddWithValue("@stl", Utility.EmptytoZero(txtShukkinTl.Text).ToString());    // 出勤日数
            sCom.Parameters.AddWithValue("@ytl", Utility.EmptytoZero(txtYukyuHiTl.Text).ToString());    // 有給日数
            sCom.Parameters.AddWithValue("@yttl", Utility.EmptytoZero(txtYukyuTmTl.Text).ToString());   // 有給時間
            sCom.Parameters.AddWithValue("@tktl", Utility.EmptytoZero(txtTokkyuTl.Text).ToString());    // 特休日数
            sCom.Parameters.AddWithValue("@frtl", Utility.EmptytoZero(txtFurikyuTl.Text).ToString());   // 振休日数
            sCom.Parameters.AddWithValue("@fdtl", Utility.EmptytoZero(txtFurideTl.Text).ToString());    // 振出日数
            sCom.Parameters.AddWithValue("@chtl", Utility.EmptytoZero(txtChisouTl.Text).ToString());    // 遅刻早退
            sCom.Parameters.AddWithValue("@kktl", Utility.EmptytoZero(txtKekkinTl.Text).ToString());    // 欠勤日数
            sCom.Parameters.AddWithValue("@zktl", "0");     // 実稼動日数
            sCom.Parameters.AddWithValue("@rTl", Utility.EmptytoZero(txtRhTl.Text).ToString());         // 労働時間（時）
            sCom.Parameters.AddWithValue("@rmTl", Utility.EmptytoZero(txtRmTl.Text).ToString());        // 労働時間（分）
            sCom.Parameters.AddWithValue("@zhTl", Utility.EmptytoZero(txtZanHTl.Text).ToString());      // 残業時間（時）
            sCom.Parameters.AddWithValue("@zhTl", Utility.EmptytoZero(txtZanMTl.Text).ToString());      // 残業時間（分）
            sCom.Parameters.AddWithValue("@shTl", Utility.EmptytoZero(txtShinyaTl.Text).ToString());    // 深夜勤務時間
            sCom.Parameters.AddWithValue("@kiTl", Utility.EmptytoZero(txtKiteiTl.Text).ToString());     // 月間規定勤務時間
            sCom.Parameters.AddWithValue("@ptsw", Utility.EmptytoZero(txtPtSouwaku.Text).ToString());   // パート労働時間総枠
            sCom.Parameters.AddWithValue("@kkknn", global.flgOn);   // 確認
            sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());
            sCom.Parameters.AddWithValue("@tatekae", Utility.EmptytoZero(txtTatekaeKin.Text.Replace(",", "")).ToString()); // 立替金
            sCom.Parameters.AddWithValue("@ryohi", Utility.EmptytoZero(txtRyohi.Text.Replace(",", "")).ToString());   // 旅費交通費
            sCom.Parameters.AddWithValue("@kinmu", Utility.EmptytoZero(txtKubun.Text.Replace(",", "")).ToString());   // 勤務先区分
            sCom.Parameters.AddWithValue("@sonota", Utility.EmptytoZero(txtSonota.Text.Replace(",", "")).ToString());   // その他支給

            sCom.Parameters.AddWithValue("@ID", sID[iX]);

            sCom.Connection = dCon.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                sCom.ExecuteNonQuery();

                // 出勤簿明細テーブル
                mySql = string.Empty;
                mySql += "update 出勤簿明細 set ";
                mySql += "休暇記号=?,有給記号=?,開始時=?,開始分=?,終了時=?,終了分=?,規定内時=?,規定内分=?,";
                mySql += "深夜帯時=?,深夜帯分=?,実働時=?,実働分=?,実働編集=?,更新年月日=? ";
                mySql += "where ID = ?";
                errMsg = "出勤簿明細テーブル更新";
                sCom.CommandText = mySql;

                for (int i = 0; i < _MULTIGYO; i++)
                {
                    sCom.Parameters.Clear();

                    sCom.Parameters.AddWithValue("@kyuka", Utility.NulltoStr(dataGridView1[cKyuka, i].Value));  // 休暇記号
                    sCom.Parameters.AddWithValue("@yukyu", Utility.NulltoStr(dataGridView1[cYukyu, i].Value));  // 有給休暇
                    sCom.Parameters.AddWithValue("@csh", Utility.NulltoStr(dataGridView1[cSH, i].Value));       // 開始時

                    // 開始分
                    if (Utility.NulltoStr(dataGridView1[cSM, i].Value) != string.Empty)
                        sCom.Parameters.AddWithValue("@csm", Utility.NulltoStr(dataGridView1[cSM, i].Value).PadLeft(2, '0'));
                    else sCom.Parameters.AddWithValue("@csm", Utility.NulltoStr(dataGridView1[cSM, i].Value));

                    sCom.Parameters.AddWithValue("@ceh", Utility.NulltoStr(dataGridView1[cEH, i].Value));       // 終了時

                    // 終了分
                    if (Utility.NulltoStr(dataGridView1[cEM, i].Value) != string.Empty)
                        sCom.Parameters.AddWithValue("@cem", Utility.NulltoStr(dataGridView1[cEM, i].Value).PadLeft(2, '0'));
                    else sCom.Parameters.AddWithValue("@cem", Utility.NulltoStr(dataGridView1[cEM, i].Value));

                    sCom.Parameters.AddWithValue("@ckkh", Utility.NulltoStr(dataGridView1[cKKH, i].Value));     // 規定内時

                    // 規定内分
                    if (Utility.NulltoStr(dataGridView1[cKKM, i].Value) != string.Empty)
                        sCom.Parameters.AddWithValue("@ckkm", Utility.NulltoStr(dataGridView1[cKKM, i].Value).PadLeft(2, '0'));
                    else sCom.Parameters.AddWithValue("@ckkm", Utility.NulltoStr(dataGridView1[cKKM, i].Value));

                    sCom.Parameters.AddWithValue("@cksh", Utility.NulltoStr(dataGridView1[cKSH, i].Value));     // 深夜帯時

                    // 深夜帯分
                    if (Utility.NulltoStr(dataGridView1[cKSM, i].Value) != string.Empty)
                        sCom.Parameters.AddWithValue("@cksm", Utility.NulltoStr(dataGridView1[cKSM, i].Value).PadLeft(2, '0'));
                    else sCom.Parameters.AddWithValue("@cksm", Utility.NulltoStr(dataGridView1[cKSM, i].Value));

                    sCom.Parameters.AddWithValue("@cth", Utility.NulltoStr(dataGridView1[cTH, i].Value));       // 実働時

                    // 実働分
                    if (Utility.NulltoStr(dataGridView1[cTM, i].Value) != string.Empty)
                        sCom.Parameters.AddWithValue("@ctm", Utility.NulltoStr(dataGridView1[cTM, i].Value).PadLeft(2, '0'));
                    else sCom.Parameters.AddWithValue("@ctm", Utility.NulltoStr(dataGridView1[cTM, i].Value));

                    // 編集チェック
                    if (dataGridView1[cCheck, i].Value.ToString() == "True")
                        sCom.Parameters.AddWithValue("@ccheck", global.flgOn);
                    else sCom.Parameters.AddWithValue("@ccheck", global.flgOff);

                    sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());  // 更新年月日
                    sCom.Parameters.AddWithValue("@ID", dataGridView1[cID, i].Value.ToString());  // ID                   

                    // テーブル書き込み
                    sCom.ExecuteNonQuery();
                }

                //トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            cI =  sID.Length - 1;
            DataShow(cI, sID, dataGridView1);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (cI > 0)
            {
                cI--;
                DataShow(cI, sID, dataGridView1);
            }   
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            cI = 0;
            DataShow(cI, sID, dataGridView1);
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     エラーチェックメイン処理 </summary>
        /// <param name="sID">
        ///     開始ID</param>
        /// <param name="eID">
        ///     終了ID</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///------------------------------------------------------------------
        private Boolean ErrCheckMain(string sIx, string eIx)
        {
            int rCnt = 0;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB();

            //エラー情報初期化
            ErrInitial();

            // 出勤簿データ読み出し
            Boolean eCheck = true;
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;

            mySql += "select * from 出勤簿ヘッダ order by ID";

            sCom.CommandText = mySql;
            sCom.Connection = dCon.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                //指定範囲のIDならエラーチェックを実施する
                if (Int64.Parse(dR["ID"].ToString()) >= Int64.Parse(sIx) && Int64.Parse(dR["ID"].ToString()) <= Int64.Parse(eIx))
                {
                    eCheck = ErrCheckData(dR);
                    if (!eCheck) break;　//エラーがあったとき
                }
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            //エラー有りの処理
            if (!eCheck)
            {
                //エラーデータのインデックスを取得
                for (int i = 0; i < sID.Length; i++)
                {
                    if (sID[i] == global.errID)
                    {
                        //エラーデータを画面表示
                        cI = i;
                        DataShow(cI, sID, dataGridView1);
                        break;
                    }
                }
            }

            return eCheck;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック </summary>
        /// <param name="cdR">
        ///     データリーダー</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///---------------------------------------------------------------------
        private Boolean ErrCheckData(OleDbDataReader cdR)
        {
            string sDate;
            DateTime eDate;

            DateTime sTime;     // 開始時刻
            DateTime eTime;     // 終了時刻

            // 未確認データ
            if (cdR["確認"].ToString() == global.flgOff.ToString())
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eNoCheck;
                global.errRow = 0;
                global.errMsg = "未確認の出勤簿です";

                return false;
            }

            //対象年
            if (Utility.NumericCheck(cdR["年"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYearMonth;
                global.errRow = 0;
                global.errMsg = "年が正しくありません";

                return false;
            }

            if (int.Parse(cdR["年"].ToString()) < 1)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYearMonth;
                global.errRow = 0;
                global.errMsg = "年が正しくありません";

                return false;
            }

            //対象月
            if (Utility.NumericCheck(cdR["月"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "月が正しくありません";

                return false;
            }

            if (int.Parse(cdR["月"].ToString()) < 1 || int.Parse(cdR["月"].ToString()) > 12)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "月が正しくありません";

                return false;
            }

            // 対象年月
            sDate = (int.Parse(cdR["年"].ToString()) + Utility.GetRekiHosei()).ToString() + "/" + cdR["月"].ToString() + "/01";
            if (DateTime.TryParse(sDate, out eDate) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYearMonth;
                global.errRow = 0;
                global.errMsg = "年月が正しくありません";

                return false;
            }

            // 勤務先区分
            if (cdR["勤務先区分"].ToString() != global.KINMU_MAIN.ToString() &&
                cdR["勤務先区分"].ToString() != global.KINMU_SUB.ToString())
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eKINMU_KUBUN;
                global.errRow = 0;
                global.errMsg = "勤務先区分（メイン勤務地：１、それ以外：２）が正しくありません";

                return false;
            }

            // 勤務先区分[1]の重複を検証する
            SysControl.SetDBConnect mdb = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            sCom.Connection = mdb.cnOpen();
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            sb.Append("select first(ID) as ID,個人番号,勤務先区分 from 出勤簿ヘッダ  ");
            sb.Append("where 勤務先区分 = ?");
            sb.Append("group by 個人番号,勤務先区分 ");
            sb.Append("having count(個人番号) > 1");
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@kk", global.KINMU_MAIN);
            sCom.CommandText = sb.ToString();

            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                global.errID = dR["ID"].ToString();
                global.errNumber = global.eKINMU_KUBUN;
                global.errRow = 0;
                global.errMsg = "メイン勤務先の出勤簿が複数あります";
                dR.Close();
                sCom.Connection.Close();
                return false;
            }

            dR.Close();

            // 勤務先区分[1]が存在しない社員番号を抽出する
            sb.Clear();
            sb.Append("select 出勤簿ヘッダ.ID, 出勤簿ヘッダ.個人番号, x.勤務先区分 from 出勤簿ヘッダ ");
            sb.Append("left join(select 個人番号, 勤務先区分 from 出勤簿ヘッダ ");
            sb.Append("where 勤務先区分 = 1) as x ");
            sb.Append("on 出勤簿ヘッダ.個人番号 = x.個人番号 ");
            sb.Append("where x.勤務先区分 is null");

            sCom.CommandText = sb.ToString();

            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                global.errID = dR["ID"].ToString();
                global.errNumber = global.eKINMU_KUBUN;
                global.errRow = 0;
                global.errMsg = "メイン勤務先の出勤簿が存在しません";
                dR.Close();
                sCom.Connection.Close();
                return false;
            }

            dR.Close();
            sCom.Connection.Close();

            // 社員番号
            // 数字以外のとき
            if (Utility.NumericCheck(cdR["個人番号"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "社員番号が入力されていません";

                return false;
            }

            // 社員番号マスター登録検査
            //dbControl.DataControl dCon = new dbControl.DataControl(_PCADBName);
            //OleDbDataReader sdR;
            //StringBuilder sb = new StringBuilder();
            //sb.Clear();
            //sb.Append("select Id from Shain where Code = '");
            //sb.Append(cdR["個人番号"].ToString() + "'");
            //sdR = dCon.FreeReader(sb.ToString());

            //bool eCnt = sdR.HasRows;
            //sdR.Close();

            //if (!eCnt)
            //{
            //    global.errID = cdR["ID"].ToString();
            //    global.errNumber = global.eShainNo;
            //    global.errRow = 0;
            //    global.errMsg = "マスター未登録の社員番号です";
            //    return false;
            //}


            if (!Utility.ComboShain.isXlsCode(xS, Utility.StrtoInt(cdR["個人番号"].ToString())))
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "マスター未登録の社員番号です";
                return false;
            }

            // 部門マスター登録検査
            //sb.Clear();
            //sb.Append("select Id from Bumon where Code = '");
            //sb.Append(cdR["所属コード"].ToString() + "'");
            //sdR = dCon.FreeReader(sb.ToString());

            //eCnt = sdR.HasRows;
            //sdR.Close();
            //dCon.Close();

            //if (!eCnt)
            //{
            //    global.errID = cdR["ID"].ToString();
            //    global.errNumber = global.eShozoku;
            //    global.errRow = 0;
            //    global.errMsg = "マスター未登録の所属コードです";
            //    return false;
            //}

            if (!Utility.ComboShain.isXlSzCode(xS, Utility.StrtoInt(cdR["所属コード"].ToString())))
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShozoku;
                global.errRow = 0;
                global.errMsg = "マスター未登録の所属コードです";
                return false;
            }

            //// 個人番号の重複を検証する
            //sCom.Connection = dCon.cnOpen();
            //StringBuilder sb = new StringBuilder();
            //sb.Clear();
            //sb.Append("select ID,個人番号 from 勤務記録ヘッダ ");
            //sb.Append("where ID <> ? and 個人番号 = ?");
            //sCom.Parameters.Clear();
            //sCom.Parameters.AddWithValue("@ID", cdR["ID"].ToString());
            //sCom.Parameters.AddWithValue("@Num", cdR["個人番号"].ToString());
            //sCom.CommandText = sb.ToString();

            //dR = sCom.ExecuteReader();

            //if (dR.HasRows)
            //{
            //    dR.Read();
            //    global.errID = cdR["ID"].ToString();
            //    global.errNumber = global.eShainNo;
            //    global.errRow = 0;
            //    global.errMsg = "個人番号が重複しています（" + dR["ID"].ToString() + "）";
            //    dR.Close();
            //    sCom.Connection.Close();
            //    return false;
            //}
            //else
            //{
            //    dR.Close();
            //    sCom.Connection.Close();
            //}

            //// 日付の重複を検証する
            //sCom.Connection = dCon.cnOpen();
            //sb.Clear();
            //sb.Append("select 日付 from 勤務記録明細 ");
            //sb.Append("where (ヘッダID = ?) and (日付 <> '') ");
            //sb.Append("group by 日付 having count(日付) <> 1");
            //sCom.Parameters.Clear();
            //sCom.CommandText = sb.ToString();
            //sCom.Parameters.AddWithValue("@HID", cdR["ID"].ToString());

            //dR = sCom.ExecuteReader();
            //string strD = string.Empty;
            //string errD = string.Empty;

            //while (dR.Read())
            //{
            //    if (strD != string.Empty) strD += "、";
            //    strD += dR["日付"].ToString() + "日";
            //    if (errD == string.Empty) errD = dR["日付"].ToString();
            //}
            //dR.Close();
            //sCom.Connection.Close();

            //// 重複日付が存在したとき
            //int eRows = 0;
            //if (strD != string.Empty)
            //{
            //    // 重複日付の位置を調べます
            //    sCom.Connection = dCon.cnOpen();
            //    sCom.CommandText = "select * from 勤務記録明細 where ヘッダID = ? order by ID";
            //    sCom.Parameters.Clear();
            //    sCom.Parameters.AddWithValue("@HID", cdR["ID"].ToString());
            //    dR = sCom.ExecuteReader();

            //    while (dR.Read())
            //    {
            //        if (dR["日付"].ToString() == errD) break;
            //        eRows++;
            //    }
            //    dR.Close();
            //    sCom.Connection.Close();

            //    global.errID = cdR["ID"].ToString();
            //    global.errNumber = global.eDay;
            //    global.errRow = eRows;
            //    global.errMsg = "日付が重複しています（" + strD +"）";
            //    return false;
            //}

            // 出勤簿明細データ
            //SysControl.SetDBConnect mdb = new SysControl.SetDBConnect();
            //OleDbCommand sCom = new OleDbCommand();
            //OleDbDataReader dR;
            sCom.Connection = mdb.cnOpen();
            sCom.CommandText = "select * from 出勤簿明細 where ヘッダID = ? order by ID";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@HID", cdR["ID"].ToString());
            dR = sCom.ExecuteReader();

            //日付別データ
            int iX = 0;
            string k = string.Empty;    // 特別休暇記号
            string yk = string.Empty;   // 有給記号

            // 集計クラス
            //sumData sDt = new sumData();

            while (dR.Read())
            {
                // 日付インデックス加算
                iX++;

                // 日付は数字か
                if (Utility.NumericCheck(dR["日付"].ToString()) == false)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eDay;
                    global.errRow = iX - 1;
                    global.errMsg = "日が正しくありません";
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                sDate = int.Parse(cdR["年"].ToString()) + Utility.GetRekiHosei() + "/" +
                        cdR["月"].ToString() + "/" + dR["日付"].ToString();

                // 存在しない日付に記入があるとき
                if (!DateTime.TryParse(sDate, out eDate))
                {
                    if (Utility.NulltoStr(dR["休暇記号"]) != string.Empty || Utility.NulltoStr(dR["有給記号"]) != string.Empty ||
                    Utility.NulltoStr(dR["開始時"]) != string.Empty || Utility.NulltoStr(dR["開始分"]) != string.Empty ||
                    Utility.NulltoStr(dR["終了時"]) != string.Empty || Utility.NulltoStr(dR["終了分"]) != string.Empty ||
                    Utility.NulltoStr(dR["規定内時"]) != string.Empty || Utility.NulltoStr(dR["規定内分"]) != string.Empty ||
                    Utility.NulltoStr(dR["深夜帯時"]) != string.Empty || Utility.NulltoStr(dR["深夜帯分"]) != string.Empty ||
                    Utility.NulltoStr(dR["実働時"]) != string.Empty || Utility.NulltoStr(dR["実働分"]) != string.Empty)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eDay;
                        global.errRow = iX - 1;
                        global.errMsg = "この行には記入できません";
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }
                }

                // 無記入の行はチェック対象外とする
                if (Utility.NulltoStr(dR["休暇記号"]) == string.Empty && Utility.NulltoStr(dR["有給記号"]) == string.Empty &&
                    Utility.NulltoStr(dR["開始時"]) == string.Empty && Utility.NulltoStr(dR["開始分"]) == string.Empty &&
                    Utility.NulltoStr(dR["終了時"]) == string.Empty && Utility.NulltoStr(dR["終了分"]) == string.Empty &&
                    Utility.NulltoStr(dR["規定内時"]) == string.Empty && Utility.NulltoStr(dR["規定内分"]) == string.Empty &&
                    Utility.NulltoStr(dR["深夜帯時"]) == string.Empty && Utility.NulltoStr(dR["深夜帯分"]) == string.Empty &&
                    Utility.NulltoStr(dR["実働時"]) == string.Empty && Utility.NulltoStr(dR["実働分"]) == string.Empty)
                {
                    continue;
                }

                // 休暇記号と有給休暇記号を取得
                k = Utility.NulltoStr(dR["休暇記号"]);
                yk = Utility.NulltoStr(dR["有給記号"]);

                // 休暇記号
                if (k != string.Empty && k != global.TOKUBETSU_KYUKA && k != global.KEKKIN_KYUKA &&
                    k != global.CHISOU_KYUKA && k != global.FURIDE_KYUKA && k != global.FURIKYU_KYUKA)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eTokubetsu;
                    global.errRow = iX - 1;
                    global.errMsg = "休暇記号が正しくありません";
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 有給休暇記号
                //if (yk != string.Empty && yk != global.ZENNICHI_YUKYU && yk != global.H1_YUKYU &&
                //    yk != global.H2_YUKYU && yk != global.H3_YUKYU && yk != global.H4_YUKYU &&
                //    yk != global.H5_YUKYU && yk != global.H6_YUKYU && yk != global.H7_YUKYU)
                //{
                //    global.errID = cdR["ID"].ToString();
                //    global.errNumber = global.eYukyu;
                //    global.errRow = iX - 1;
                //    global.errMsg = "有給記号が正しくありません";
                //    dR.Close();
                //    sCom.Connection.Close();
                //    return false;
                //}

                // 有給休暇記号 2016/11/17
                if (yk != string.Empty && yk != global.ZENNICHI_YUKYU && yk != global.HANNICHI_YUKYU)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eYukyu;
                    global.errRow = iX - 1;
                    global.errMsg = "有給記号が正しくありません";
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 開始時間・時チェック
                if (!errCheckTime(dR["開始時"], cdR, "開始時間", k, yk, iX, "H", global.eSH))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 開始時間・分チェック
                if (!errCheckTime(dR["開始分"], cdR, "開始時間", k, yk, iX, "M", global.eSM))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 終了時間・時チェック
                if (!errCheckTime(dR["終了時"], cdR, "終了時間", k, yk, iX, "H", global.eEH))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 終了時間・分チェック
                if (!errCheckTime(dR["終了分"], cdR, "終了時間", k, yk, iX, "M", global.eEM))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 終了時刻範囲
                if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24 &&
                    Utility.StrtoInt(Utility.NulltoStr(dR["終了分"])) > 0)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eEM;
                    global.errRow = iX - 1;
                    global.errMsg = "終了時刻範囲を超えています（～２４：００）";
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }


                // 規定内休憩・時間チェック
                if (!errCheckKyukeiTime(dR["規定内時"], cdR, "規定内休憩時間", k, yk, iX, "H", global.eKKH))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 規定内休憩・分チェック
                if (!errCheckKyukeiTime(dR["規定内分"], cdR, "規定内休憩時間", k, yk, iX, "M", global.eKKM))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 深夜帯休憩・時間チェック
                if (!errCheckKyukeiTime(dR["深夜帯時"], cdR, "深夜帯休憩時間", k, yk, iX, "H", global.eKSH))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 深夜帯休憩・分チェック
                if (!errCheckKyukeiTime(dR["深夜帯分"], cdR, "規定内休憩時間", k, yk, iX, "M", global.eKSM))
                {
                    dR.Close();
                    sCom.Connection.Close();
                    return false;
                }

                // 開始時刻・終了時刻チェック
                string sh = string.Empty;
                string eh = string.Empty;

                if (Utility.NulltoStr(dR["開始時"]) != string.Empty &&
                    Utility.NulltoStr(dR["開始分"]) != string.Empty &&
                    Utility.NulltoStr(dR["終了時"]) != string.Empty &&
                    Utility.NulltoStr(dR["終了分"]) != string.Empty)
                {
                    // 開始時刻取得
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) == 24)
                        sTime = DateTime.Parse("0:" + Utility.NulltoStr(dR["開始分"]));
                    else sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));

                    // 終了時刻取得
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24)
                        eTime = DateTime.Parse("23:59");
                    else eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));

                    //sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));
                    //eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));

                    // 開始時刻 > 終了時刻のときNG
                    if (DateTime.Compare(sTime, eTime) > 0)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eEH;
                        global.errRow = iX - 1;
                        global.errMsg = "終了時刻が開始時刻以前になっています";
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }

                    // 開始時刻～終了時刻と休憩時間
                    double w = 0;

                    // 2013/07/03 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24 &&
                        Utility.StrtoInt(Utility.NulltoStr(dR["終了分"])) == 0)
                        w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes + 1;
                    else w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;  // 稼働時間

                    double kk = Utility.StrtoInt(Utility.NulltoStr(dR["規定内時"])) * 60 + Utility.StrtoInt(Utility.NulltoStr(dR["規定内分"]));
                    double ks = Utility.StrtoInt(Utility.NulltoStr(dR["深夜帯時"])) * 60 + Utility.StrtoInt(Utility.NulltoStr(dR["深夜帯分"]));
                    if (w < (kk + ks))
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eKKH;
                        global.errRow = iX - 1;
                        global.errMsg = "稼働時間より休憩時間が長くなっています";
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }

                    // 規定内稼働時間 2013/07/03 ※終了時間が５時以降を対象とする
                    w = 0;
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) > 5)
                    {
                        if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) < 5 ||
                            Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) == 24)
                            sTime = DateTime.Parse("05:00");
                        else sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));

                        if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) >= 22)
                            eTime = DateTime.Parse("22:00");
                        else eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));

                        // 規定内稼働時間
                        w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }

                    // 規定内休憩時間のチェック
                    kk = Utility.StrtoInt(Utility.NulltoStr(dR["規定内時"])) * 60 + Utility.StrtoInt(Utility.NulltoStr(dR["規定内分"]));
                    if (w < kk)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eKKH;
                        global.errRow = iX - 1;
                        global.errMsg = "規定内稼働時間より規定内休憩時間が長くなっています";
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }

                    // 深夜帯休憩時間のチェック
                    w = 0;

                    // 早朝時間帯：終了時間 2013/07/03
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) >= 5)
                        eTime = global.dt0500;
                    else eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));

                    //////if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) == 24)
                    //////{
                    //////    sTime = DateTime.Parse("00:" + Utility.NulltoStr(dR["開始分"]));
                    //////    w = Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;  // 深夜帯稼働時間
                    //////}
                    //////else if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) < 5)
                    //////{
                    //////    sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));
                    //////    w = Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;  // 深夜帯稼働時間
                    //////}

                    // 早朝時間帯：開始時間 2013/07/03
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) == 24)
                    {
                        sTime = DateTime.Parse("00:" + Utility.NulltoStr(dR["開始分"]));
                        w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;  // 深夜帯稼働時間
                    }
                    else if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) < 5)
                    {
                        sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));
                        w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;  // 深夜帯稼働時間
                    }

                    // 22:00 以降時間帯 2013/07/03
                    if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) >= 22)
                    {
                        if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) < 24)
                            eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));
                        else eTime = DateTime.Parse("23:59");

                        w += Utility.GetTimeSpan(global.dt2200, eTime).TotalMinutes;  // 深夜帯稼働時間

                        // 2013/07/03 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
                        if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24 &&
                            Utility.StrtoInt(Utility.NulltoStr(dR["終了分"])) == 0)
                            w += 1;
                    }

                    ks = Utility.StrtoInt(Utility.NulltoStr(dR["深夜帯時"])) * 60 + Utility.StrtoInt(Utility.NulltoStr(dR["深夜帯分"]));
                    if (w < ks)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eKSH;
                        global.errRow = iX - 1;
                        global.errMsg = "深夜稼働時間より深夜休憩時間が長くなっています";
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }

                    // 実働編集フラグが0のとき実働時間チェックを行う
                    if (dR["実働編集"].ToString() == "0")
                    {
                        // 開始時刻取得
                        if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) == 24)
                            sTime = DateTime.Parse("0:" + Utility.NulltoStr(dR["開始分"]));
                        else sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));

                        // 終了時刻取得
                        if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24)
                            eTime = DateTime.Parse("23:59");
                        else eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));

                        if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24 &&
                            Utility.StrtoInt(Utility.NulltoStr(dR["終了分"])) == 0)
                            w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes + 1;  // 終了時間が24:00記入のときは23:59まで計算して稼働時間1分加算する
                        else w = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;

                        kk = Utility.StrtoInt(Utility.NulltoStr(dR["規定内時"])) * 60 + Utility.StrtoInt(Utility.NulltoStr(dR["規定内分"]));
                        ks = Utility.StrtoInt(Utility.NulltoStr(dR["深夜帯時"])) * 60 + Utility.StrtoInt(Utility.NulltoStr(dR["深夜帯分"]));

                        double zw = w - kk - ks;    // 実稼働時間計算
                        int zh = (int)(System.Math.Floor(zw / 60));
                        int zm = (int)(zw % 60);

                        // 記入値と比較
                        if (zh != Utility.StrtoInt(Utility.NulltoStr(dR["実働時"])) ||
                            zm != Utility.StrtoInt(Utility.NulltoStr(dR["実働分"])))
                        {
                            global.errID = cdR["ID"].ToString();
                            global.errNumber = global.eTH;
                            global.errRow = iX - 1;
                            global.errMsg = "実働時間が正しくありません（" + zh.ToString() + "時間 " + zm.ToString() + "分）";
                            dR.Close();
                            sCom.Connection.Close();
                            return false;
                        }
                    }
                }
                else
                {
                    // 開始終了時刻が無記入で稼働時間が記入されているとき
                    if (Utility.NulltoStr(dR["実働時"]) != string.Empty ||
                        Utility.NulltoStr(dR["実働分"]) != string.Empty)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eTH;
                        global.errRow = iX - 1;
                        global.errMsg = "開始終了時刻が無記入で実働時間が記入されています";
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }
                }

                //// 集計処理
                //SumDataCal(dR, sDt);
            }

            // 出勤簿明細データリーダークローズ
            dR.Close();
            sCom.Connection.Close();

            //// 出勤簿ヘッダデータ更新
            //SumDataUpdate(cdR["ID"].ToString());

            return true;
        }


        ///--------------------------------------------------------------------
        /// <summary>
        ///     同日勤務エラーチェックメイン処理 </summary>
        /// <param name="sID">
        ///     開始ID</param>
        /// <param name="eID">
        ///     終了ID</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///--------------------------------------------------------------------
        private Boolean ErrCheckSameTime()
        {
            string hID = string.Empty;
            int rCnt = 0;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB();

            //エラー情報初期化
            ErrInitial();

            // 出勤簿データ読み出し
            Boolean eCheck = true;
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;

            mySql += "select 出勤簿ヘッダ.ID, 出勤簿ヘッダ.個人番号, 出勤簿明細.ヘッダID, 出勤簿ヘッダ.氏名, ";
            mySql += "出勤簿明細.日付, 出勤簿明細.開始時, 出勤簿明細.開始分, 出勤簿明細.終了時, 出勤簿明細.終了分 ";
            mySql += "from 出勤簿ヘッダ inner join 出勤簿明細 ";
            mySql += "on 出勤簿ヘッダ.ID = 出勤簿明細.ヘッダID ";
            mySql += "order by 出勤簿ヘッダ.ID, 出勤簿明細.ID";

            sCom.CommandText = mySql;
            sCom.Connection = dCon.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                // データ件数加算
                if (hID != dR["ID"].ToString())
                {
                    rCnt++;
                    hID = dR["ID"].ToString();
                }

                // プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                // エラーチェック
                if (Utility.NulltoStr(dR["開始時"].ToString()) != string.Empty)
                    eCheck = SameTimeOtherPlaceErrCheck(dR);

                if (!eCheck) break;　//エラーがあったとき
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            //エラー有りの処理
            if (!eCheck)
            {
                //エラーデータのインデックスを取得
                for (int i = 0; i < sID.Length; i++)
                {
                    if (sID[i] == global.errID)
                    {
                        //エラーデータを画面表示
                        cI = i;
                        DataShow(cI, sID, dataGridView1);
                        break;
                    }
                }
            }

            return eCheck;
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     同じ社員が別勤務場所で同日勤務したときのチェック </summary>
        /// <param name="dR">
        ///     出勤簿データリーダー</param>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///--------------------------------------------------------------------
        private bool SameTimeOtherPlaceErrCheck(OleDbDataReader dR)
        {
            OleDbCommand sCom = new OleDbCommand();
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader r;

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select 出勤簿ヘッダ.所属名,出勤簿明細.* from 出勤簿ヘッダ inner join 出勤簿明細 ");
            sb.Append("on 出勤簿ヘッダ.ID = 出勤簿明細.ヘッダID ");
            sb.Append("where 出勤簿ヘッダ.ID <> '" + dR["ID"].ToString() + "' and ");
            sb.Append("出勤簿ヘッダ.個人番号 = '" + dR["個人番号"].ToString() + "' and ");
            sb.Append("出勤簿明細.日付 = '" + dR["日付"].ToString() + "' and ");
            sb.Append("出勤簿明細.開始時 <> ''");

            sCom.CommandText = sb.ToString();
            r = sCom.ExecuteReader();

            // 同じ日付の出勤簿データがあるか？
            if (!r.HasRows)
            {
                r.Close();
                sCom.Connection.Close();
                return true;
            }

            DateTime sTime = DateTime.Parse("0:00");
            DateTime eTime = DateTime.Parse("0:00");
            DateTime sTime2 = DateTime.Parse("0:00");
            DateTime eTime2 = DateTime.Parse("0:00");
            string sName = string.Empty;
            string sShozoku = string.Empty;

            // 開始時刻取得
            if (Utility.StrtoInt(Utility.NulltoStr(dR["開始時"])) == 24)
                sTime = DateTime.Parse("0:" + Utility.NulltoStr(dR["開始分"]));
            else sTime = DateTime.Parse(Utility.NulltoStr(dR["開始時"]) + ":" + Utility.NulltoStr(dR["開始分"]));

            // 終了時刻取得
            if (Utility.StrtoInt(Utility.NulltoStr(dR["終了時"])) == 24)
                eTime = DateTime.Parse("23:59");
            else eTime = DateTime.Parse(Utility.NulltoStr(dR["終了時"]) + ":" + Utility.NulltoStr(dR["終了分"]));
            
            // 別出勤簿の同日時刻取得              
            while (r.Read())
            {
                // 氏名取得
                sName = dR["氏名"].ToString();

                // 所属取得
                sShozoku = Utility.NulltoStr(r["所属名"]);

                // 別出勤簿の同日開始時刻取得
                if (Utility.StrtoInt(Utility.NulltoStr(r["開始時"])) == 24)
                    sTime2 = DateTime.Parse("0:" + Utility.NulltoStr(r["開始分"]));
                else sTime2 = DateTime.Parse(Utility.NulltoStr(r["開始時"]) + ":" + Utility.NulltoStr(r["開始分"]));

                // 別出勤簿の同日終了時刻取得
                if (Utility.StrtoInt(Utility.NulltoStr(r["終了時"])) == 24)
                    eTime2 = DateTime.Parse("23:59");
                else eTime2 = DateTime.Parse(Utility.NulltoStr(r["終了時"]) + ":" + Utility.NulltoStr(r["終了分"]));
           
            }
            r.Close();
            sCom.Connection.Close();

            bool ec = true;

            // １．勤務終了時刻が勤務時間帯である
            if (DateTime.Compare(eTime2, sTime) > 0 && DateTime.Compare(eTime2, eTime) <= 0) ec = false;

            // ２．勤務開始時刻が勤務時間帯である
            if (DateTime.Compare(sTime2, sTime) >= 0 && DateTime.Compare(sTime2, eTime) < 0) ec = false;

            // ３．勤務開始時刻と終了時刻が同じである
            if (DateTime.Compare(sTime2, sTime) == 0 && DateTime.Compare(eTime2, eTime) == 0) ec = false;

            // ４．勤務開始時刻～終了時刻が勤務時間帯である
            if (DateTime.Compare(sTime2, sTime) > 0 && DateTime.Compare(eTime2, eTime) < 0) ec = false;

            // NGのとき
            if (!ec)
            {
                global.errID = dR["ID"].ToString();
                global.errNumber = global.eSH;
                global.errRow = int.Parse(dR["日付"].ToString()) - 1;
                global.errMsg = "勤務時間帯が" + sShozoku + "の同日勤務時間と重複しています";
                return false;
            }
            else return true;
        }

        /// <summary>
        /// 出勤簿ヘッダテーブル更新
        /// </summary>
        /// <param name="hID">ID</param>
        private void SumDataUpdate(string hID)
        {
            //MDB接続
            SysControl.SetDBConnect mdb = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = mdb.cnOpen();

            // 出勤簿ヘッダテーブル
            string mySql = string.Empty;
            mySql += "update 出勤簿ヘッダ set ";
            mySql += "出勤日数合計=?,有休日数合計=?,有休時間合計=?,特休日数合計=?,欠勤日数合計=?,";
            mySql += "実稼動日数合計=?,総労働=?,総労働分=?,残業時=?,残業分=?,深夜勤務時間合計=?,";
            mySql += "月間規定勤務時間=?,更新年月日=?";
            mySql += "where ID = ?";
            
            sCom.CommandText = mySql;
            sCom.Parameters.AddWithValue("@stl", Utility.EmptytoZero(txtShukkinTl.Text).ToString());    // 出勤日数
            sCom.Parameters.AddWithValue("@ytl", Utility.EmptytoZero(txtYukyuHiTl.Text).ToString());    // 有給日数
            sCom.Parameters.AddWithValue("@yttl", Utility.EmptytoZero(txtYukyuTmTl.Text).ToString());   // 有給時間
            sCom.Parameters.AddWithValue("@tktl", Utility.EmptytoZero(txtTokkyuTl.Text).ToString());    // 特休日数
            sCom.Parameters.AddWithValue("@kktl", Utility.EmptytoZero(txtKekkinTl.Text).ToString());    // 欠勤日数
            sCom.Parameters.AddWithValue("@zktl", "0");     // 実稼動日数
            sCom.Parameters.AddWithValue("@rTl", Utility.EmptytoZero(txtRhTl.Text).ToString());         // 労働時間（時）
            sCom.Parameters.AddWithValue("@rmTl", Utility.EmptytoZero(txtRmTl.Text).ToString());        // 労働時間（分）
            sCom.Parameters.AddWithValue("@zhTl", Utility.EmptytoZero(txtZanHTl.Text).ToString());      // 残業時間（時）
            sCom.Parameters.AddWithValue("@zhTl", Utility.EmptytoZero(txtZanMTl.Text).ToString());      // 残業時間（分）
            sCom.Parameters.AddWithValue("@shTl", Utility.EmptytoZero(txtShinyaTl.Text).ToString());    // 深夜勤務時間
            sCom.Parameters.AddWithValue("@kiTl", Utility.EmptytoZero(txtKiteiTl.Text).ToString());    // 月間規定勤務時間
            sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());

            sCom.Parameters.AddWithValue("@ID", hID);

            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }


        /// <summary>
        /// エラーチェックボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnErrCheck_Click(object sender, EventArgs e)
        {
            //カレントレコード更新
            CurDataUpDate(cI);

            //エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(sID[cI], sID[sID.Length - 1]) == false) return;

            //エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (cI > 0)
            {
                if (ErrCheckMain(sID[0], sID[cI - 1]) == false) return;
            }

            // 同日勤務エラーチェック
            if (!ErrCheckSameTime()) return;

            MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dataGridView1.CurrentCell = null;

        }
        
        /// <summary>
        /// エラー表示
        /// </summary>
        private void ErrShow()
        {
            if (global.errNumber != global.eNothing)
            {
                lblErrMsg.Visible = true;
                lblErrMsg.Text = global.errMsg;

                // 勤務先区分
                if (global.errNumber == global.eKINMU_KUBUN)
                {
                    txtKubun.BackColor = Color.Yellow;
                    txtKubun.Focus();
                }

                // 対象年月
                if (global.errNumber == global.eYearMonth)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtMonth.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                // 対象月
                if (global.errNumber == global.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // 所属コード
                if (global.errNumber == global.eShozoku)
                {
                    txtShozokuCode.BackColor = Color.Yellow;
                    txtShozokuCode.Focus();
                }

                // 個人番号
                if (global.errNumber == global.eShainNo)
                {
                    txtNo.BackColor = Color.Yellow;
                    txtNo.Focus();
                }

                // 日
                if (global.errNumber == global.eDay)
                {
                    dataGridView1[cDay, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cDay, global.errRow];
                }

                // 特別休暇
                if (global.errNumber == global.eTokubetsu)
                {
                    dataGridView1[cKyuka, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cKyuka, global.errRow];
                }

                // 有給休暇
                if (global.errNumber == global.eYukyu)
                {
                    dataGridView1[cYukyu, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cYukyu, global.errRow];
                }

                // 開始時
                if (global.errNumber == global.eSH)
                {
                    dataGridView1[cSH, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cSH, global.errRow];
                }

                // 開始分
                if (global.errNumber == global.eSM)
                {
                    dataGridView1[cSM, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cSM, global.errRow];
                }

                // 終了時
                if (global.errNumber == global.eEH)
                {
                    dataGridView1[cEH, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cEH, global.errRow];
                }

                // 終了分
                if (global.errNumber == global.eEM)
                {
                    dataGridView1[cEM, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cEM, global.errRow];
                }

                // 規定内時
                if (global.errNumber == global.eKKH)
                {
                    dataGridView1[cKKH, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cKKH, global.errRow];
                }

                // 規定内分
                if (global.errNumber == global.eKKM)
                {
                    dataGridView1[cKKM, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cKKM, global.errRow];
                }

                // 深夜帯時
                if (global.errNumber == global.eKSH)
                {
                    dataGridView1[cKSH, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cKSH, global.errRow];
                }

                // 深夜帯分
                if (global.errNumber == global.eKSM)
                {
                    dataGridView1[cKSM, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cKSM, global.errRow];
                }

                // 実働時
                if (global.errNumber == global.eTH)
                {
                    dataGridView1[cTH, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cTH, global.errRow];
                }

                // 実働分
                if (global.errNumber == global.eTM)
                {
                    dataGridView1[cTM, global.errRow].Style.BackColor = Color.Yellow;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[cTM, global.errRow];
                }
            }
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            cI = hScrollBar1.Value;
            DataShow(cI, sID, dataGridView1);
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の出勤簿データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // レコードと画像ファイルを削除する
            DataDelete();

            //テーブル件数カウント：ゼロならばプログラム終了
            if (CountMDB() == 0)
            {
                MessageBox.Show("全ての出勤簿データが削除されました。処理を終了します。", "出勤簿削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //終了処理
                Environment.Exit(0);
            }

            //テーブルデータキー項目読み込み
            sID = LoadMdbID();

            //エラー情報初期化
            ErrInitial();

            //レコードを表示
            if (sID.Length - 1 < cI) cI = sID.Length - 1;
            DataShow(cI, sID, dataGridView1);
        }

        private void DataDelete()
        {
            //カレントデータを削除します
            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();

            // 画像ファイル名を取得します
            string sImgNm = string.Empty;
            OleDbDataReader dR;
            sCom.CommandText = "select 画像名 from 出勤簿ヘッダ where ID = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@ID", sID[cI]);
            dR = sCom.ExecuteReader();
            while (dR.Read())
            {
                sImgNm = dR["画像名"].ToString();
            }
            dR.Close();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //勤務記録ヘッダデータを削除します
                sCom.CommandText = "delete from 出勤簿ヘッダ where ID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID[cI]);
                sCom.ExecuteNonQuery();

                //勤務記録明細データを削除します
                sCom.CommandText = "delete from 出勤簿明細 where ヘッダID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID[cI]);
                sCom.ExecuteNonQuery();

                //画像ファイルを削除する
                if (System.IO.File.Exists(Properties.Settings.Default.dataPath + sImgNm))
                {
                    System.IO.File.Delete(Properties.Settings.Default.dataPath + sImgNm);
                }

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("出勤簿の削除に失敗しました" + Environment.NewLine + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Tag.ToString() != END_MAKEDATA)
            {
                if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                // カレントデータ更新
                if (dID == string.Empty) CurDataUpDate(cI);
            }

            // 解放する
            this.Dispose();
        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("給与計算用勤怠データを作成します。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // カレントレコード更新
            CurDataUpDate(cI);

            // エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(sID[cI], sID[sID.Length - 1]) == false) return;

            // エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (cI > 0)
            {
                if (ErrCheckMain(sID[0], sID[cI - 1]) == false) return;
            }

            // 同日勤務エラーチェック
            if (!ErrCheckSameTime()) return;

            // 汎用データ作成
            SaveDataJcs();
        }
        
        ///------------------------------------------------------
        /// <summary>
        ///     常陽コンピュータサービス向け勤怠データ作成 </summary>
        ///------------------------------------------------------
        private void SaveDataJcs()
        {
            // 出力データ生成
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            try
            {
                //オーナーフォームを無効にする
                this.Enabled = false;

                // ヘッダID 2014/10/28
                string hdID = string.Empty;

                // 個人番号
                string wsID = string.Empty;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                //レコード件数取得
                int cTotal = CountMDBitem();
                int rCnt = 1;

                // データベース接続
                sCom.Connection = Con.cnOpen();

                string pyymm = string.Empty;

                // 出勤簿ヘッダデータリーダーを取得します
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("SELECT 出勤簿ヘッダ.* from 出勤簿ヘッダ ");
                sb.Append("order by 個人番号,勤務先区分,ID ");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                ////出力先フォルダがあるか？なければ作成する
                if (!System.IO.Directory.Exists(Properties.Settings.Default.okPath))
                    System.IO.Directory.CreateDirectory(Properties.Settings.Default.okPath);
                
                // 明細書き出し
                sumData sd = null;
                //double pZan = 0;

                // 集計結果配列
                string[] outXls = null;
                int iX = 0;

                while (dR.Read())
                {
                    //プログレスバー表示
                    frmP.Text = "汎用データ作成中です・・・" + rCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = rCnt / cTotal * 100;
                    frmP.ProgressStep();

                    // 社員ＩＤでブレーク発生
                    if (wsID != string.Empty && wsID != dR["個人番号"].ToString())
                    {
                        // データを配列へ出力
                        sd.SaveDataJcs(ref outXls, iX, xS);
                        iX++;
                    }

                    // 合計クラス
                    if (wsID != dR["個人番号"].ToString())
                    {
                        sd = new sumData();
                        sd.cCnt = 0;
                    }

                    // 社員毎に集計
                    sd.CaltotalJcs(dR);

                    // 社員ID
                    wsID = dR["個人番号"].ToString();

                    // ヘッダID
                    hdID = dR["ID"].ToString();
                }

                // データを配列へ出力
                sd.SaveDataJcs(ref outXls, iX, xS);

                // ＣＳＶデータへ追加出力 2016/12/06
                Utility.csvFileWrite(Properties.Settings.Default.instPath + _gDir + @"\" + Properties.Settings.Default.outCsvName, outXls, "");
                
                // データリーダーをクローズ
                dR.Close();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 画像ファイル退避
                tifFileMove();

                // 過去データ作成
                SaveLastData();

                // 出勤簿ヘッダレコード削除
                sCom.CommandText = "delete from 出勤簿ヘッダ";
                sCom.ExecuteNonQuery();

                // 出勤簿明細レコード削除
                sCom.CommandText = "delete from 出勤簿明細";
                sCom.ExecuteNonQuery();

                ////設定月数分経過した過去画像およびデータを削除する
                imageDelete();

                //終了
                MessageBox.Show("給与ＣＳＶデータ作成が終了しました", "給与計算用ＣＳＶデータ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();

                //MDBファイル最適化
                mdbCompact();
            }
        }

        ///------------------------------------------------------
        /// <summary>
        ///     ＰＣＡ給与勤怠データ作成 </summary>
        ///     
        /// *********　2013年ヴァージョン・現在は不使用 **********
        ///------------------------------------------------------
        private void SaveData()
        {
            // 出力データ生成
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            try
            {
                //オーナーフォームを無効にする
                this.Enabled = false;

                // ヘッダID 2014/10/28
                string hdID = string.Empty;

                // 社員ID
                string wsID = string.Empty;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                //レコード件数取得
                int cTotal = CountMDBitem();
                int rCnt = 1;

                // データベース接続
                sCom.Connection = Con.cnOpen();

                string pyymm = string.Empty;

                // 出勤簿ヘッダデータリーダーを取得します
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("SELECT 出勤簿ヘッダ.* from 出勤簿ヘッダ ");
                sb.Append("order by 出勤簿ヘッダ.社員ID,出勤簿ヘッダ.ID ");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                ////出力先フォルダがあるか？なければ作成する
                if (!System.IO.Directory.Exists(Properties.Settings.Default.okPath))
                    System.IO.Directory.CreateDirectory(Properties.Settings.Default.okPath);

                //出力ファイルインスタンス作成
                string iFile = global.OKFILE + "_";
                iFile += DateTime.Today.Year.ToString() + string.Format("{0:00}", DateTime.Today.Month) + string.Format("{0:00}", DateTime.Today.Day);
                iFile += string.Format("{0:00}", DateTime.Now.Hour) + string.Format("{0:00}", DateTime.Now.Minute) + string.Format("{0:00}", DateTime.Now.Second);
                iFile += ".dat";

                StreamWriter outFile = new StreamWriter(Properties.Settings.Default.okPath + iFile, false, System.Text.Encoding.GetEncoding(932));

                // 明細書き出し
                sumData sd = null;
                double pZan = 0;

                while (dR.Read())
                {
                    //プログレスバー表示
                    frmP.Text = "汎用データ作成中です・・・" + rCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = rCnt / cTotal * 100;
                    frmP.ProgressStep();

                    // 社員ＩＤでブレーク発生
                    if (wsID != string.Empty && wsID != dR["社員ID"].ToString())
                    {
                        // 複数勤務票のとき 2014/10/28
                        if (sd.cCnt > 1)
                        {
                            // 残業時間・社員、一部パートの月間規定勤務時間設定者 2014/10/24
                            if (sd.kitei != 0)
                            {
                                pZan = getZangyoTimeTotal(wsID, sd.kitei);
                                sd.ZangyoH = (int)System.Math.Floor(pZan / 60);
                                sd.ZangyoM = (int)(pZan % 60);
                            }
                            else if (sd.KyuyoKbn == "1")   // パートタイマー残業時間計算 2014/10/24
                            {
                                pZan = getZangyoPartTotal(sd.cYear, sd.cMonth, sd.SouroudouH.ToString(), sd.SouroudouM.ToString());
                                sd.ZangyoH = (int)System.Math.Floor(pZan / 60);
                                sd.ZangyoM = (int)(pZan % 60);
                            }

                            ////// パートタイマー残業時間計算 2014/10/24
                            ////if (sd.KyuyoKbn == "1")
                            ////{
                            ////    pZan = getZangyoPartTotal(sd.cYear, sd.cMonth, sd.SouroudouH.ToString(), sd.SouroudouM.ToString());

                            ////    sd.ZangyoH = (int)System.Math.Floor(pZan / 60);
                            ////    sd.ZangyoM = (int)(pZan % 60);
                            ////}
                        }

                        // 汎用データの出力
                        sd.SaveDatacsv(outFile);
                    }

                    // 合計クラス
                    if (wsID != dR["社員ID"].ToString())
                    {
                        sd = new sumData();
                        sd.cCnt = 0;
                    }

                    // 社員毎に集計
                    sd.Caltotal(dR);

                    // 社員ID
                    wsID = dR["社員ID"].ToString();

                    // ヘッダID
                    hdID = dR["ID"].ToString();
                }
                
                // 複数勤務票のとき 2014/10/28
                if (sd.cCnt > 1)
                {
                    // 残業時間・社員、一部パートの月間規定勤務時間設定者 2014/10/24
                    if (sd.kitei != 0)
                    {
                        pZan = getZangyoTimeTotal(wsID, sd.kitei);
                        sd.ZangyoH = (int)System.Math.Floor(pZan / 60);
                        sd.ZangyoM = (int)(pZan % 60);
                    }
                    else if (sd.KyuyoKbn == "1")   // パートタイマー残業時間計算 2014/10/24
                    {
                        pZan = getZangyoPartTotal(sd.cYear, sd.cMonth, sd.SouroudouH.ToString(), sd.SouroudouM.ToString());
                        sd.ZangyoH = (int)System.Math.Floor(pZan / 60);
                        sd.ZangyoM = (int)(pZan % 60);
                    }

                    //// パートタイマー残業時間計算 2014/10/24
                    //if (sd.KyuyoKbn == "1")
                    //{
                    //    pZan = getZangyoPartTotal(sd.cYear, sd.cMonth, sd.SouroudouH.ToString(), sd.SouroudouM.ToString());

                    //    sd.ZangyoH = (int)System.Math.Floor(pZan / 60);
                    //    sd.ZangyoM = (int)(pZan % 60);
                    //}
                }

                // 汎用データの出力
                sd.SaveDatacsv(outFile);

                // データリーダーをクローズ
                dR.Close();

                // 出力ファイルをクローズ
                outFile.Close();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 画像ファイル退避
                tifFileMove();

                // 過去データ作成
                SaveLastData();

                // 出勤簿ヘッダレコード削除
                sCom.CommandText = "delete from 出勤簿ヘッダ";
                sCom.ExecuteNonQuery();

                // 出勤簿明細レコード削除
                sCom.CommandText = "delete from 出勤簿明細";
                sCom.ExecuteNonQuery();

                ////設定月数分経過した過去画像およびデータを削除する
                imageDelete();

                //終了
                MessageBox.Show("終了しました。PCA給与Xでデータの受け入れを行ってください。", "給与計算用勤怠データ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();

                //MDBファイル最適化
                mdbCompact();
            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     画像ファイル退避処理 </summary>
        ///-----------------------------------------------------------------
        private void tifFileMove()
        {
            // DocuWorksレジストリ・キーを取得します
            string rKeyName = @"SOFTWARE\FujiXerox\MPM3\SystemInfo";
            Microsoft.Win32.RegistryKey rKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rKeyName);
            
            // ローカルmdb接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbCommand sCom2 = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom2.Connection = Con.cnOpen();
            OleDbDataReader dR;

            // 移動先フォルダがあるか？なければ作成する（TIFフォルダ）
            if (!System.IO.Directory.Exists(Properties.Settings.Default.tifPath))
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.tifPath);

            // docuworkフォルダがあるか？なければ作成する（docuworkフォルダ）
            if (!System.IO.Directory.Exists(Properties.Settings.Default.docuworkPath))
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.docuworkPath);

            // 出勤簿ヘッダのデータリーダーを取得する
            sCom.CommandText = "select * from 出勤簿ヘッダ order by 個人番号";
            dR = sCom.ExecuteReader();

            string kNum = string.Empty;
            //int seqNum = 0;
            while (dR.Read())
            {
                //if (kNum == dR["個人番号"].ToString())
                //{
                //    seqNum++;
                //}
                //else
                //{
                //    seqNum = 0;
                //}

                string NewFilenameYearMonth = (int.Parse(dR["年"].ToString()) + Utility.GetRekiHosei()).ToString() + 
                                              dR["月"].ToString().PadLeft(2, '0');

                // 画像ファイルパスを取得する
                string fromImg = Properties.Settings.Default.dataPath + dR["画像名"].ToString();

                //
                //  tifファイル処理
                //

                // ファイル名を「対象年月個人番号 + 部門コード」に変えて退避先フォルダへ移動する
                string NewFileNumber = dR["個人番号"].ToString().PadLeft(5, '0') + dR["所属コード"].ToString().PadLeft(5, '0');
                string NewFilename = NewFilenameYearMonth + NewFileNumber + ".tif";
                string toImg = Properties.Settings.Default.tifPath + NewFilename;

                // 同名ファイルが既に登録済みのときは削除する
                if (System.IO.File.Exists(toImg)) System.IO.File.Delete(toImg);

                // ファイルを移動する
                File.Move(fromImg, toImg);

                // 出勤簿ヘッダレコードの画像ファイル名を書き換える
                sCom2.CommandText = "update 出勤簿ヘッダ set 画像名=? where ID=?";
                sCom2.Parameters.Clear();
                sCom2.Parameters.AddWithValue("@img", NewFilename);
                sCom2.Parameters.AddWithValue("@ID", dR["ID"].ToString());
                sCom2.ExecuteNonQuery();

                //
                //  docuworkファイル処理
                //
                //  DocuWorksのレジストリが登録されているとき実行する
                //

                if (rKey != null)
                {
                    // docuwork出力ファイルパスを定義
                    string outPath = string.Empty;
                    outPath = Properties.Settings.Default.docuworkPath + NewFilename.Replace(".tif", ".xdw");

                    // 同名docuworkファイルが既に登録済みのときは削除する
                    if (System.IO.File.Exists(outPath)) System.IO.File.Delete(outPath);

                    // docuworkファイル出力
                    xdwCreate.FromTiff(toImg, outPath);
                }

                // 個人番号
                kNum = dR["個人番号"].ToString();
            }

            dR.Close();
            sCom.Connection.Close();
            sCom2.Connection.Close();
        }

        /// <summary>
        /// 受渡データ出力
        /// </summary>
        /// <param name="outFile">出力するStreamWriterオブジェクト</param>
        /// <param name="sd">集計データクラス</param>
        private void SaveDatacsv(StreamWriter outFile, sumData sd)
        {
            // CSVファイルを書き出す
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append(sd.c1 + sd.c2 + sd.c3 + sd.c4 + sd.c5 + sd.c6 + sd.c7 + sd.c8 + sd.c9 + sd.c10);
            sb.Append(sd.c11 + sd.c12 + sd.c13 + sd.c14 + sd.c15 + sd.c16 + sd.c17 + sd.c18 + sd.c19 + sd.c20);
            sb.Append(sd.c21 + sd.c22 + sd.c23 + sd.c24 + sd.c25 + sd.c26 + sd.c27 + sd.c28 + sd.c29 + sd.c30);
            sb.Append(sd.c31 + sd.c32 + sd.c33 + sd.c34 + sd.c35 + sd.c36 + sd.c37 + sd.c38 + sd.c39 + sd.c40);

            for (int i = 0; i < sd.c41.Length; i++)
            {
                sb.Append(sd.c41[i]);
                
            }

            //明細ファイル出力
            outFile.WriteLine(sb.ToString());
        }

        /// <summary>
        /// MDBファイルを最適化する
        /// </summary>
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();
                string OldDb = Properties.Settings.Default.mdbOlePath;
                string NewDb = Properties.Settings.Default.mdbPathTemp;

                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.mdbFileBack);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.mdbFile, Properties.Settings.Default.mdbFileBack);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.mdbFileTemp, Properties.Settings.Default.mdbFile);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
        
        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        /// <summary>
        /// 設定月数分経過した過去画像を削除する    
        /// </summary>
        private void imageDelete()
        {
            //削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (Properties.Settings.Default.imageDeleteSpan == global.flgOff) return;

            try
            {
                //削除年月の取得
                DateTime dt = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");  
                DateTime delDate = dt.AddMonths(Properties.Settings.Default.imageDeleteSpan * (-1));
                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月
                //int _waYYMM = (delDate.Year - Properties.Settings.Default.RekiHosei) * 100 + _dMM;   //基準年月(和暦）
                int _DataYYMM;
                string fileYYMM;

                //設定月数分経過した過去画像を削除する            
                foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.tifPath, "*.tif"))
                {
                    // ファイル名が規定外のファイルは読み飛ばします
                    if (System.IO.Path.GetFileName(files).Length < 24) continue;

                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(18, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }

                //設定月数分経過したdocuworks画像を削除する            
                foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.docuworkPath, "*.xdw"))
                {
                    // ファイル名が規定外のファイルは読み飛ばします
                    if (System.IO.Path.GetFileName(files).Length < 24) continue;

                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(18, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }

                // 過去データを削除する
                SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                OleDbCommand sCom2 = new OleDbCommand();
                OleDbConnection Cn = Con.cnOpen();
                sCom.Connection = Cn;
                sCom2.Connection = Cn;
                OleDbDataReader dR = null;

                // 基準年月以前の過去データを取得します
                sCom.CommandText = "select * from 過去出勤簿ヘッダ where 年*100+月 <= ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@YearMonth", _dYYMM.ToString());
                dR = sCom.ExecuteReader();

                // 過去出勤簿明細レコードを削除します
                while (dR.Read())
                {
                    sCom2.CommandText = "delete from 過去出勤簿明細 where ヘッダID=? ";
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@ID", dR["ID"].ToString());
                    sCom2.ExecuteNonQuery();
                }
                dR.Close();

                // 過去出勤簿ヘッダレコードを削除します
                sCom.CommandText = "delete from 過去出勤簿ヘッダ where 年*100+月 <= ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@YearMonth", _dYYMM.ToString());
                sCom.ExecuteNonQuery();

                Cn.Close();
                sCom.Connection.Close();
                sCom2.Connection.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //選択画面表示
            this.Hide();
            frmShainSel frm = new frmShainSel();
            frm.ShowDialog();
            string selID = frm._ID;
            frm.Dispose();
            this.Show();

            //勤務票が選択されていないときは終了
            if (selID == string.Empty) return;

            //カレントデータの更新
            CurDataUpDate(cI);

            //エラー情報初期化
            ErrInitial();

            //選択されたレコードへ移動する
            for (int i = 0; i < sID.Length; i++)
            {
                if (sID[i] == selID)
                {
                    cI = i;                             //カレントレコードindexをセット
                    DataShow(cI, sID, dataGridView1);    //データ表示
                    break;
                }
            }
        }
        
        /// <summary>
        /// 過去データ登録
        /// </summary>
        private void SaveLastData()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbCommand sCom2 = new OleDbCommand();
            OleDbCommand sCom3 = new OleDbCommand();
            OleDbConnection Cn = Con.cnOpen();
            sCom.Connection = Cn;
            sCom2.Connection = Cn;
            sCom3.Connection = Cn;
            OleDbDataReader dR = null;
            OleDbDataReader dR2 = null;
            StringBuilder sb = new StringBuilder();

            try
            {
                // 同じ年月、個人番号の過去データを削除します
                sb.Clear();
                sb.Append("select * from 出勤簿ヘッダ order by ID");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();
                while (dR.Read())
                {
                    sCom2.CommandText = "select * from 過去出勤簿ヘッダ where 個人番号=? and 年=? and 月=?";
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@ID", dR["個人番号"].ToString());
                    sCom2.Parameters.AddWithValue("@Year", int.Parse(dR["年"].ToString()) + Utility.GetRekiHosei());
                    sCom2.Parameters.AddWithValue("@Month", dR["月"].ToString());
                    dR2 = sCom2.ExecuteReader();

                    // 過去出勤簿明細レコードを削除します
                    while (dR2.Read())
                    {
                        sCom3.CommandText = "delete from 過去出勤簿明細 where ヘッダID=? ";
                        sCom3.Parameters.Clear();
                        sCom3.Parameters.AddWithValue("@ID", dR2["ID"].ToString());
                        sCom3.ExecuteNonQuery();
                    }
                    dR2.Close();

                    // 過去出勤簿ヘッダレコードを削除します
                    sCom3.CommandText = "delete from 過去出勤簿ヘッダ where 個人番号=? and 年=? and 月=?";
                    sCom3.Parameters.Clear();
                    sCom3.Parameters.AddWithValue("@ID", dR["個人番号"].ToString());
                    sCom3.Parameters.AddWithValue("@Year", int.Parse(dR["年"].ToString()) + Utility.GetRekiHosei());
                    sCom3.Parameters.AddWithValue("@Month", dR["月"].ToString());
                    sCom3.ExecuteNonQuery();
                }
                dR.Close();

                // 過去出勤簿ヘッダレコードを作成します
                sb.Clear();
                sb.Append("insert into 過去出勤簿ヘッダ ");
                sb.Append("select * from 出勤簿ヘッダ ");
                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();

                // 和暦の過去出勤簿ヘッダの年を西暦に変換
                sb.Clear();
                sb.Append("update 過去出勤簿ヘッダ ");
                sb.Append("set 年 = 年 + " + Utility.GetRekiHosei().ToString());
                sb.Append(" where 年 < 100 ");
                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();

                // 過去出勤簿ヘッダにデータ領域名をセットします
                sb.Clear();
                sb.Append("update 過去出勤簿ヘッダ ");
                sb.Append("set データ領域名 = '" + "" + "' ");
                //sb.Append("set データ領域名 = '" + _PCAComName + "' ");
                sb.Append("where データ領域名 = ''");
                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();

                // 出勤簿明細のデータリーダーを取得します
                sb.Clear();
                sb.Append("select * from 出勤簿明細 order by ID");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                // 過去出勤簿明細レコード作成用SQL文定義
                sb.Clear();
                sb.Append("insert into 過去出勤簿明細 (");
                sb.Append("ヘッダID,日付,休暇記号,有給記号,開始時,開始分,終了時,終了分,規定内時,規定内分,");
                sb.Append("深夜帯時,深夜帯分,実働時,実働分,実働編集,更新年月日) ");
                sb.Append("values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

                sCom2.CommandText = sb.ToString();

                // 過去出勤簿明細レコードを作成します
                string _hd = string.Empty;
                while (dR.Read())
                {
                    // ヘッダIDの検証
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@HDID", dR["ヘッダID"].ToString());
                    sCom2.Parameters.AddWithValue("@DAY", dR["日付"].ToString());
                    sCom2.Parameters.AddWithValue("@KYU", dR["休暇記号"].ToString());
                    sCom2.Parameters.AddWithValue("@YU", dR["有給記号"].ToString());
                    sCom2.Parameters.AddWithValue("@T8", dR["開始時"].ToString());
                    sCom2.Parameters.AddWithValue("@T9", dR["開始分"].ToString());
                    sCom2.Parameters.AddWithValue("@T10", dR["終了時"].ToString());
                    sCom2.Parameters.AddWithValue("@T11", dR["終了分"].ToString());
                    sCom2.Parameters.AddWithValue("@T12", dR["規定内時"].ToString());
                    sCom2.Parameters.AddWithValue("@T13", dR["規定内分"].ToString());
                    sCom2.Parameters.AddWithValue("@T14", dR["深夜帯時"].ToString());
                    sCom2.Parameters.AddWithValue("@T15", dR["深夜帯分"].ToString());
                    sCom2.Parameters.AddWithValue("@T16", dR["実働時"].ToString());
                    sCom2.Parameters.AddWithValue("@T17", dR["実働分"].ToString());
                    sCom2.Parameters.AddWithValue("@T18", dR["実働編集"].ToString());
                    sCom2.Parameters.AddWithValue("@UPDAY", DateTime.Today.ToShortDateString());
                    sCom2.ExecuteNonQuery();
                }
                dR.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "過去出勤簿ヘッダ作成エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                if (sCom2.Connection.State == ConnectionState.Open) sCom2.Connection.Close();
                if (sCom3.Connection.State == ConnectionState.Open) sCom3.Connection.Close(); 
                if (Cn.State == ConnectionState.Open) Cn.Close();
            }
        }

        /// <summary>
        /// 休日配列作成
        /// </summary>
        private void GetHolidayArray()
        {
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
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //if (e.RowIndex < 0) return;

            string colName = dataGridView1.Columns[e.ColumnIndex].Name;

            if (colName == cSH || colName == cSE || colName == cEH || colName == cEE ||
                colName == cKKH || colName == cKKE || colName == cKSH || colName == cKSE ||
                colName == cTH || colName == cTE)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name;
            if (colName == cCheck)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    dataGridView1.RefreshEdit();
                }
            }
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellEnter_1(object sender, DataGridViewCellEventArgs e)
        {
            string ColH = string.Empty;
            string ColM = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name;

            // 開始時間または終了時間を判断
            if (ColM == cSM)    // 開始時刻
            {
                ColH = cSH;
            }
            else if (ColM == cEM)   // 終了時刻
            {
                ColH = cEH;
            }
            else if (ColM == cKKM)  // 規定内
            {
                ColH = cKKH;
            }
            else if (ColM == cKSM)  // 深夜帯
            {
                ColH = cKSH;
            }
            else
            {
                return;
            }

            // 開始時、終了時が入力済みで開始分、終了分が未入力のとき"00"を表示します
            if (dataGridView1[ColH, dataGridView1.CurrentRow.Index].Value != null)
            {
                if (dataGridView1[ColH, dataGridView1.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
                {
                    if (dataGridView1[ColM, dataGridView1.CurrentRow.Index].Value == null)
                    {
                        dataGridView1[ColM, dataGridView1.CurrentRow.Index].Value = "00";
                    }
                    else if (dataGridView1[ColM, dataGridView1.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
                    {
                        dataGridView1[ColM, dataGridView1.CurrentRow.Index].Value = "00";
                    }
                }
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     時間記入チェック </summary>
        /// <param name="obj">
        ///     データーリーダー項目オブジェクト</param>
        /// <param name="cdR">
        ///     出勤簿データリーダーオブジェクト</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="k">
        ///     特別休暇</param>
        /// <param name="yk">
        ///     有給休暇</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="Mode">
        ///     時間：H, 分:M</param>
        /// <param name="errNum">
        ///     エラー箇所番号</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///-----------------------------------------------------------------------
        private bool errCheckTime(object obj, OleDbDataReader cdR, string tittle, string k, string yk, int iX, string Mode, int errNum)
        {
            bool rtn = true;

            // 無記入のとき
            if (Utility.NulltoStr(obj) == string.Empty)
            {
                // 社員：特別休暇・欠勤・振休・終日有給以外で無記入のときNGとする：2016/11/17
                if (cdR["給与区分"].ToString() == global.STATUS_SHAIN.ToString())
                {
                    if (k != global.TOKUBETSU_KYUKA && k != global.KEKKIN_KYUKA && 
                        k != global.FURIKYU_KYUKA && yk != global.ZENNICHI_YUKYU)
                    {
                        global.errMsg = tittle + "が未入力です";
                        rtn = false;
                    }
                }
                else // パートタイマー：欠勤・振休以外で無記入のときNGとする
                {
                    if (k != global.KEKKIN_KYUKA && k != global.FURIKYU_KYUKA)
                    {
                        global.errMsg = tittle + "が未入力です";
                        rtn = false;
                    }
                }
            }
            else
            {
                // 欠勤で記入されているときNGとする
                if (k == global.KEKKIN_KYUKA)
                {
                    global.errMsg = "欠勤で" + tittle + "が入力されています";
                    rtn = false;
                }

                // 振休で記入されているときNGとする
                if (k == global.FURIKYU_KYUKA)
                {
                    global.errMsg = "振休で" + tittle + "が入力されています";
                    rtn = false;
                }

                // 社員で特休または有給で記入されているときNGとする：2016/11/17
                if (cdR["給与区分"].ToString() == global.STATUS_SHAIN.ToString())
                {
                    // 特別休暇で記入されているときNGとする
                    if (k == global.TOKUBETSU_KYUKA)
                    {
                        global.errMsg = "特別休暇で" + tittle + "が入力されています";
                        rtn = false;
                    }

                    if (yk == global.ZENNICHI_YUKYU)
                    {
                        global.errMsg = "終日有給で" + tittle + "が入力されています";
                        rtn = false;
                    }
                }

                // 数字以外の記入
                if (!Utility.NumericCheck(obj.ToString()))
                {
                    global.errMsg = tittle + "が正しくありません";
                    rtn = false;
                }

                if (Mode == "H")
                {
                    if (int.Parse(obj.ToString()) < 0 || int.Parse(obj.ToString()) > 24)
                    {
                        global.errMsg = tittle + "が正しくありません";
                        rtn = false;
                    }
                }
                else if (Mode == "M")
                {
                    if (int.Parse(obj.ToString()) < 0 || int.Parse(obj.ToString()) > 59)
                    {
                        global.errMsg = tittle + "が正しくありません";
                        rtn = false;
                    }
                }
            }

            // 戻り値
            if (!rtn)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = errNum;
                global.errRow = iX - 1;
                return false;
            }
            else return true;
        }

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     休憩時間記入チェック <summary>
        /// <param name="obj">
        ///     データーリーダー項目オブジェクト</param>
        /// <param name="cdR">
        ///     出勤簿データリーダーオブジェクト</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="k">
        ///     特別休暇</param>
        /// <param name="yk">
        ///     有給休暇</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="Mode">
        ///     時間：H, 分:M</param>
        /// <param name="errNum">
        ///     エラー箇所番号</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///---------------------------------------------------------------------------
        private bool errCheckKyukeiTime(object obj, OleDbDataReader cdR, string tittle, string k, string yk, int iX, string Mode, int errNum)
        {
            // 無記入のとき
            if (Utility.NulltoStr(obj) != string.Empty)
            {
                // 欠勤で記入されているときNGとする
                if (k == global.KEKKIN_KYUKA)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = errNum;
                    global.errRow = iX - 1;
                    global.errMsg = "欠勤で" + tittle + "が入力されています";
                    return false;
                }

                // 振休で記入されているときNGとする
                if (k == global.FURIKYU_KYUKA)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = errNum;
                    global.errRow = iX - 1;
                    global.errMsg = "振休で" + tittle + "が入力されています";
                    return false;
                }

                // 社員で特休または有給で記入されているときNGとする : 2016/11/17
                if (cdR["給与区分"].ToString() == global.STATUS_SHAIN.ToString())
                {
                    // 特別休暇で記入されているときNGとする
                    if (k == global.TOKUBETSU_KYUKA)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = errNum;
                        global.errRow = iX - 1;
                        global.errMsg = "特別休暇で" + tittle + "が入力されています";
                        return false;
                    }

                    if (yk == global.ZENNICHI_YUKYU)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = errNum;
                        global.errRow = iX - 1;
                        global.errMsg = "終日有給で" + tittle + "が入力されています";
                        return false;
                    }
                }

                // 数字以外の記入
                if (!Utility.NumericCheck(obj.ToString()))
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = errNum;
                    global.errRow = iX - 1;
                    global.errMsg = tittle + "が正しくありません";
                    return false;
                }

                if (Mode == "H")
                {
                    if (int.Parse(obj.ToString()) < 0 || int.Parse(obj.ToString()) > 24)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = errNum;
                        global.errRow = iX - 1;
                        global.errMsg = tittle + "が正しくありません";
                        return false;
                    }
                }
                else if (Mode == "M")
                {
                    if (int.Parse(obj.ToString()) < 0 || int.Parse(obj.ToString()) > 59)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = errNum;
                        global.errRow = iX - 1;
                        global.errMsg = tittle + "が正しくありません";
                        return false;
                    }
                }
            }
            return true;
        }

        private void txtShozokuCode_TextChanged(object sender, EventArgs e)
        {
            if (dID != string.Empty)
            {
                return;
            } 

            this.lblShozoku.Text = Utility.ComboShain.getXlSzName(xS, Utility.StrtoInt(txtShozokuCode.Text));

            //// SQLServer接続
            //dbControl.DataControl dCon = new dbControl.DataControl(_PCADBName);
            //OleDbDataReader dR;

            //// 部門データリーダーを取得する
            //StringBuilder sb = new StringBuilder();
            //sb.Append("select Bumon.Name from Bumon ");
            //sb.Append("where Bumon.Code = '" + txtShozokuCode.Text.Trim().PadLeft(3, '0') + "'");

            //dR = dCon.FreeReader(sb.ToString());

            //while (dR.Read())
            //{
            //    this.lblShozoku.Text = dR["Name"].ToString().Trim();
            //}

            //dR.Close();
            //dCon.Close();

        }
        
        ///-------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        ///-------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                global.pblImageFile = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    if (Properties.Settings.Default.imgDpi == 300)
                    {
                        leadImg.ScaleFactor *= global.ZOOM_RATE;
                    }
                    else if (Properties.Settings.Default.imgDpi == 200)
                    {
                        leadImg.ScaleFactor *= global.ZOOM_RATE200;
                    }
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                global.pblImageFile = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
            }
        }

        private void leadImg_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void leadImg_MouseMove(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("実働時間編集チェックを全てオンにします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[cCheck, i].Value = true;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("実働時間編集チェックを全てオフにします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[cCheck, i].Value = false;
            }
        }

    }
}
