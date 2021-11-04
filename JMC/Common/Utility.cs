using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace JMC.Common
{
    class Utility
    {
        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// 休日コンボボックスクラス
        /// </summary>
        public class comboHoliday
        {
            public string Date { get; set; }
            public string Name { get; set; }

            /// <summary>
            /// 休日コンボボックスデータロード
            /// </summary>
            /// <param name="tempBox">ロード先コンボボックスオブジェクト名</param>
            public static void Load(ComboBox tempBox)
            {

                // 休日配列
                string[] sDay = {"01/01元旦", "     成人の日", "02/11建国記念の日", "     春分の日", "04/29昭和の日",
                            "05/03憲法記念日","05/04みどりの日","05/05こどもの日","     海の日","     敬老の日",
                            "     秋分の日","     体育の日","11/03文化の日","11/23勤労感謝の日","12/23天皇誕生日",
                            "     振替休日","     国民の休日","     土曜日","     年末年始休暇","     夏季休暇"}; 

                try
                {
                    comboHoliday cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Date";

                    foreach (var a in sDay)
                    {
                        cmb1 = new comboHoliday();
                        cmb1.Date = a.Substring(0, 5);
                        int s = a.Length;
                        cmb1.Name = a.Substring(5, s - 5);
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "休日コンボボックスロード");
                }
            }

            /// <summary>
            /// 休日コンボ表示
            /// </summary>
            /// <param name="tempBox">コンボボックスオブジェクト</param>
            /// <param name="dt">月日</param>
            public static void selectedIndex(ComboBox tempBox, string dt)
            {
                comboHoliday cmbS = new comboHoliday();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboHoliday)tempBox.SelectedItem;

                    if (cmbS.Date == dt)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        /// <summary>
        /// 文字列の値が数字かチェックする
        /// </summary>
        /// <param name="tempStr">検証する文字列</param>
        /// <returns>数字:true,数字でない:false</returns>
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }
        
        /// <summary>
        /// emptyを"0"に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        /// <summary>
        /// Nullをstring.Empty("")に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        /// <summary>
        /// Nullをstring.Empty("")に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// <summary>
        /// 文字型をIntへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>Int型の値</returns>
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr))
            {
                return int.Parse(tempStr);
            }
            else return 0;
        }

        /// <summary>
        /// 文字型をDoubleへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>double型の値</returns>
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        /// <summary>
        /// 経過時間を返す
        /// </summary>
        /// <param name="s">開始時間</param>
        /// <param name="e">終了時間</param>
        /// <returns>経過時間</returns>
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てします。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }


        // 部門コンボボックスクラス
        public class ComboBumon
        {
            public int ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //部門マスターロード
            public static void load(ComboBox tempObj, int tempLen, string dbName)
            {
                try
                {
                    ComboBumon cmb1;
                    string sqlSTRING = string.Empty;
                    dbControl.DataControl dCon = new dbControl.DataControl(dbName);
                    OleDbDataReader dR;

                    sqlSTRING += "select * from Bumon inner join ";
                    sqlSTRING += "(select distinct BumonId as bumonid from Shain) as sbumon ";
                    sqlSTRING += "on Bumon.Id = sbumon.bumonid ";
                    sqlSTRING += "order by Code";

                    //データリーダーを取得する
                    dR = dCon.FreeReader(sqlSTRING);

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    while (dR.Read())
                    {
                        cmb1 = new ComboBumon();
                        cmb1.ID = int.Parse(dR["Id"].ToString());
                        cmb1.DisplayName = string.Format("{0:D" + tempLen.ToString() + "}", int.Parse(dR["Code"].ToString())) + " " + dR["Name"].ToString().Trim() + "";
                        cmb1.Name = dR["Name"].ToString().Trim() + "";
                        cmb1.code = dR["Code"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    dCon.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "部門コンボボックスロード");
                }
            }

            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから部門コンボボックスにロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void loadCsv(ComboBox tempObj, string fName)
            {
                string[] bArray = null;

                try
                {
                    ComboBumon cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 社員名簿CSV読み込み
                    bArray = System.IO.File.ReadAllLines(fName, Encoding.Default);
                    
                    System.Collections.ArrayList al = new System.Collections.ArrayList();

                    foreach (var t in bArray)
                    {
                        string[] d = t.Split(',');

                        if (d.Length < 4)
                        {
                            continue;
                        }

                        string bn = d[3].PadLeft(5, '0') + "," + d[2] + "";
                        al.Add(bn);
                    }

                    // 配列をソートします
                    al.Sort();

                    string alCode = string.Empty;

                    foreach (var item in al)
                    {
                        string [] d = item.ToString().Split(',');

                        // 重複コード・部門はネグる
                        if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
                        {
                            continue;
                        }

                        // コンボボックスにセット
                        cmb1 = new ComboBumon();
                        cmb1.ID = 0;
                        cmb1.DisplayName = item.ToString().Replace(',', ' ');

                        string[] cn = item.ToString().Split(',');
                        cmb1.Name = cn[1] + "";
                        cmb1.code = cn[0] + "";
                        tempObj.Items.Add(cmb1);

                        alCode = item.ToString();
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "部門コンボボックスロード");
                }
            }
        }

        // 社員コンボボックスクラス
        public class ComboShain
        {
            public int ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }
            public int YakushokuType { get; set; }
            public string BumonName { get; set; }
            public string BumonCode { get; set; }

            // 社員マスターロード
            public static void load(ComboBox tempObj, string dbName)
            {
                try
                {
                    ComboShain cmb1;
                    string sqlSTRING = string.Empty;
                    dbControl.DataControl dCon = new dbControl.DataControl(dbName);
                    OleDbDataReader dR;

                    sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
                    sqlSTRING += "where Shurojokyo = 1 ";
                    sqlSTRING += "order by Code";

                    //データリーダーを取得する
                    dR = dCon.FreeReader(sqlSTRING);

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    while (dR.Read())
                    {
                        cmb1 = new ComboShain();
                        cmb1.ID = int.Parse(dR["Id"].ToString());
                        cmb1.DisplayName = dR["Code"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
                        cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
                        cmb1.code = (dR["Code"].ToString() + "").Trim();
                        cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    dCon.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "社員コンボボックスロード");
                }

            }


            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから社員コンボボックスにロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void loadCsv(ComboBox tempObj, string fName)
            {
                string[] bArray = null;

                try
                {
                    ComboShain cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 社員名簿CSV読み込み
                    bArray = System.IO.File.ReadAllLines(fName, Encoding.Default);

                    System.Collections.ArrayList al = new System.Collections.ArrayList();

                    foreach (var t in bArray)
                    {
                        string[] d = t.Split(',');

                        if (d.Length < 5)
                        {
                            continue;
                        }

                        string bn = d[1].PadLeft(5, '0') + "," + d[0] + "";
                        al.Add(bn);
                    }

                    // 配列をソートします
                    al.Sort();

                    string alCode = string.Empty;

                    foreach (var item in al)
                    {
                        string[] d = item.ToString().Split(',');

                        // 重複社員はネグる
                        if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
                        {
                            continue;
                        }

                        // コンボボックスにセット
                        cmb1 = new ComboShain();
                        cmb1.ID = 0;
                        cmb1.DisplayName = item.ToString().Replace(',', ' ');

                        string[] cn = item.ToString().Split(',');
                        cmb1.Name = cn[1] + "";
                        cmb1.code = cn[0] + "";
                        tempObj.Items.Add(cmb1);

                        alCode = item.ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "社員コンボボックスロード");
                }

            }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     常陽コンピュータサービスエクセル社員マスターコンボボックスロード </summary>
            /// <param name="fName">
            ///     エクセルファイル名</param>
            /// <param name="sheetNum">
            ///     シート名</param>
            /// <param name="tempObj">
            ///     コンボボックス</param>
            /// <param name="szStatus">
            ///     ０：部門情報含めない、１：部門情報含める</param>
            ///------------------------------------------------------------------------
            public static void xlsArrayLoad(string fName, string sheetNum, ComboBox tempObj, int szStatus)
            {
                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(fName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[sheetNum];

                Excel.Range dRg;
                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                const int C_BCODE = 7;
                const int C_BNAME = 8;
                const int C_SCODE = 11;
                const int C_SEI = 24;
                const int C_MEI = 25;

                int iX = 0;

                System.Collections.ArrayList al = new System.Collections.ArrayList();

                try
                {
                    int frmRow = 21;  // 開始行
                    int toRow = oxlsSheet.UsedRange.Rows.Count;

                    for (int i = frmRow; i <= toRow; i++)
                    {
                        // 社員番号
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_SCODE];

                        // 社員番号に有効値があること
                        string sc = dRg.Text.ToString().Trim();
                        if (Utility.StrtoInt(sc) == 0)
                        {
                            continue;
                        }

                        // 社員姓
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_SEI];
                        string sei = dRg.Text.ToString().Trim();

                        // 社員名
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_MEI];
                        string mei = dRg.Text.ToString().Trim();

                        string bn = sc.ToString().PadLeft(5, '0') + "," + (sei + " " + mei) + "";

                        if (szStatus == global.flgOn)
                        {
                            // 組織コード
                            dRg = (Excel.Range)oxlsSheet.Cells[i, C_BCODE];
                            int bCode = Utility.StrtoInt(dRg.Text.ToString().Trim());

                            // 組織名称
                            dRg = (Excel.Range)oxlsSheet.Cells[i, C_BNAME];
                            string bName = dRg.Text.ToString().Trim();

                            bn += "," + bCode.ToString() + "," + (bName + "");

                        }

                        al.Add(bn);

                        iX++;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エクセル社員マスター読み込み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
                    oXls = null;
                    oXlsBook = null;
                    oxlsSheet = null;
                    GC.Collect();
                }

                ComboShain cmb1;

                tempObj.Items.Clear();
                tempObj.DisplayMember = "DisplayName";
                tempObj.ValueMember = "code";

                // 配列をソートします
                al.Sort();

                string alCode = string.Empty;

                foreach (var item in al)
                {
                    string[] d = item.ToString().Split(',');

                    // 重複社員はネグる
                    if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
                    {
                        continue;
                    }

                    // コンボボックスにセット
                    cmb1 = new ComboShain();
                    cmb1.ID = 0;
                    cmb1.DisplayName = item.ToString().Replace(',', ' ');

                    string[] cn = item.ToString().Split(',');
                    cmb1.Name = cn[1] + "";
                    cmb1.code = cn[0] + "";

                    if (szStatus == global.flgOn)
                    {
                        cmb1.BumonCode = cn[2] + "";
                        cmb1.BumonName = cn[3] + "";
                    }

                    tempObj.Items.Add(cmb1);

                    alCode = item.ToString();
                }
            }


            ///------------------------------------------------------------------------
            /// <summary>
            ///     CSV社員マスターコンボボックスロード </summary>
            /// <param name="fName">
            ///     CSVファイル名</param>
            /// <param name="tempObj">
            ///     コンボボックス</param>
            /// <param name="szStatus">
            ///     ０：部門情報含めない、１：部門情報含める</param>
            ///------------------------------------------------------------------------
            public static void csvArrayLoad(string fName, ComboBox tempObj, int szStatus)
            {
                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;
                
                int iX = 0;

                System.Collections.ArrayList al = new System.Collections.ArrayList();

                string[] bArray = null;

                try
                {
                    // 社員名簿CSV読み込み
                    bArray = System.IO.File.ReadAllLines(fName, Encoding.Default);

                    foreach (var t in bArray)
                    {
                        string[] d = t.Split(',');

                        if (d.Length < 5)
                        {
                            continue;
                        }

                        string bn = d[1].PadLeft(5, '0') + "," + d[0] + "";


                        if (szStatus == global.flgOn)
                        {
                            // 組織コード
                            int bCode = Utility.StrtoInt(d[3]);

                            // 組織名称
                            string bName = Utility.NulltoStr(d[2]).Trim();

                            bn += "," + bCode.ToString() + "," + (bName + "");
                        }

                        al.Add(bn);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "CSV社員マスター読み込み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                }

                ComboShain cmb1;

                tempObj.Items.Clear();
                tempObj.DisplayMember = "DisplayName";
                tempObj.ValueMember = "code";

                // 配列をソートします
                al.Sort();

                string alCode = string.Empty;

                foreach (var item in al)
                {
                    string[] d = item.ToString().Split(',');

                    // 重複社員はネグる
                    if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
                    {
                        continue;
                    }

                    // コンボボックスにセット
                    cmb1 = new ComboShain();
                    cmb1.ID = 0;
                    cmb1.DisplayName = item.ToString().Replace(',', ' ');

                    string[] cn = item.ToString().Split(',');
                    cmb1.Name = cn[1] + "";
                    cmb1.code = cn[0] + "";

                    if (szStatus == global.flgOn)
                    {
                        cmb1.BumonCode = cn[2] + "";
                        cmb1.BumonName = cn[3] + "";
                    }

                    tempObj.Items.Add(cmb1);

                    alCode = item.ToString();
                }
            }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     常陽コンピュータサービスCSV社員マスター読み込み </summary>
            /// <param name="fName">
            ///     CSVファイル名</param>
            /// <param name="xS">
            ///     読み込む配列</param>
            ///------------------------------------------------------------------------
            public static void csvArrayLoad(string fName, ref xlsShain[] xS)
            {
                // 社員名簿CSV読み込み
                string[] csvArray = System.IO.File.ReadAllLines(fName, Encoding.Default);

                int iX = 0;

                foreach (var t in csvArray)
                {
                    string[] f = t.Split(',');

                    if (f.Length < 5)
                    {
                        continue;
                    }
                    
                    // 配列を加算
                    Array.Resize(ref xS, iX + 1);
                    xS[iX] = new xlsShain();

                    // 社員番号
                    xS[iX].sCode = Utility.StrtoInt(f[1]);

                    // 組織コード
                    xS[iX].bCode = Utility.StrtoInt(f[3]);

                    // 組織名称
                    xS[iX].bName = Utility.NulltoStr(f[2]);

                    // 社員名
                    xS[iX].sName = Utility.NulltoStr(f[0]);

                    // サブ・メイン区分
                    xS[iX].kbn = Utility.StrtoInt(f[4]);

                    iX++;
                }
            }


            ///------------------------------------------------------------------------
            /// <summary>
            ///     常陽コンピュータサービスエクセル社員マスター読み込み </summary>
            /// <param name="fName">
            ///     エクセルファイル名</param>
            /// <param name="sheetNum">
            ///     シート名</param>
            /// <param name="xS">
            ///     読み込む配列</param>
            ///------------------------------------------------------------------------
            public static void xlsArrayLoad(string fName, string sheetNum, ref xlsShain [] xS)
            {
                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(fName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[sheetNum];

                Excel.Range dRg;
                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                xS = null;
                const int C_BCODE = 7;
                const int C_BNAME = 8;
                const int C_SCODE = 11;
                const int C_SEI = 24;
                const int C_MEI = 25;

                int iX = 0;

                try
                {
                    int frmRow = 21;  // 開始行
                    int toRow = oxlsSheet.UsedRange.Rows.Count;

                    for (int i = frmRow; i <= toRow; i++)
                    {
                        // 社員番号
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_SCODE];

                        // 社員番号に有効値があること
                        string sc = dRg.Text.ToString().Trim();
                        if (Utility.StrtoInt(sc) == 0)
                        {
                            continue;
                        }

                        // 配列を加算
                        Array.Resize(ref xS, iX + 1);
                        xS[iX] = new xlsShain();

                        // 社員番号
                        xS[iX].sCode = Utility.StrtoInt(sc);

                        // 組織コード
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_BCODE];
                        xS[iX].bCode = Utility.StrtoInt(dRg.Text.ToString().Trim());

                        // 組織名称
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_BNAME];
                        xS[iX].bName = dRg.Text.ToString().Trim();

                        // 社員姓
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_SEI];
                        string sei = dRg.Text.ToString().Trim();

                        // 社員名
                        dRg = (Excel.Range)oxlsSheet.Cells[i, C_MEI];
                        string mei = dRg.Text.ToString().Trim();

                        xS[iX].sName = sei + " " + mei;

                        iX++;
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "エクセル社員マスター読み込み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
                }

            }

            ///------------------------------------------------------------------
            /// <summary>
            ///     社員名簿配列から社員情報を取得する </summary>
            /// <param name="x">
            ///     社員名簿配列</param>
            /// <param name="sCode">
            ///     社員番号</param>
            /// <param name="zCode">
            ///     勤務先コード</param>
            /// <param name="zName">
            ///     勤務先名</param>
            /// <returns>
            ///     社員名</returns>
            ///------------------------------------------------------------------
            public static string getXlsSname(xlsShain[] x, int sCode, out string zCode, out string zName)
            {
                string rVal = string.Empty;
                zCode = string.Empty;
                zName = string.Empty;

                foreach (var t in x.Where(a => a.sCode == sCode))
                {
                    rVal = t.sName;
                    zCode = t.bCode.ToString();
                    zName = t.bName.ToString();
                }

                return rVal;
            }

            ///------------------------------------------------------------------
            /// <summary>
            ///     社員名簿配列から部門名を取得する </summary>
            /// <param name="x">
            ///     社員名簿配列</param>
            /// <param name="zCode">
            ///     部門コード</param>
            /// <returns>
            ///     部門名</returns>
            ///------------------------------------------------------------------
            public static string getXlSzName(xlsShain[] x, int zCode)
            {
                string rVal = string.Empty;

                foreach (var t in x.Where(a => a.bCode == zCode))
                {
                    rVal = t.bName;
                }

                return rVal;
            }
            ///------------------------------------------------------------------
            /// <summary>
            ///     社員名簿配列に指定の社員番号が存在するか調べる </summary>
            /// <param name="x">
            ///     社員名簿配列</param>
            /// <param name="sCode">
            ///     社員番号</param>
            /// <returns>
            ///     true:あり, false:なし</returns>
            ///------------------------------------------------------------------

            public static bool isXlsCode(xlsShain[] x, int sCode)
            {
                bool rVal = false;

                foreach (var t in x.Where(a => a.sCode == sCode))
                {
                    rVal = true;
                }

                return rVal;
            }

            ///------------------------------------------------------------------
            /// <summary>
            ///     社員名簿配列に指定の部門コードが存在するか調べる </summary>
            /// <param name="x">
            ///     社員名簿配列</param>
            /// <param name="sCode">
            ///     部門コード</param>
            /// <returns>
            ///     true:あり, false:なし</returns>
            ///------------------------------------------------------------------
            public static bool isXlSzCode(xlsShain[] x, int zCode)
            {
                bool rVal = false;

                foreach (var t in x.Where(a => a.bCode == zCode))
                {
                    rVal = true;
                }

                return rVal;
            }

            // パートタイマーロード
            public static void loadPart(ComboBox tempObj, string dbName)
            {
                try
                {
                    ComboShain cmb1;
                    string sqlSTRING = string.Empty;
                    dbControl.DataControl dCon = new dbControl.DataControl(dbName);
                    OleDbDataReader dR;
                    sqlSTRING += "select Bumon.Code as bumoncode,Bumon.Name as bumonname,Shain.Id as shainid,";
                    sqlSTRING += "Shain.Code as shaincode,Shain.Sei,Shain.Mei, Shain.YakushokuType ";
                    sqlSTRING += "from Shain left join Bumon ";
                    sqlSTRING += "on Shain.BumonId = Bumon.Id ";
                    sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
                    sqlSTRING += "order by Shain.Code";
                    
                    //sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
                    //sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
                    //sqlSTRING += "order by Code";

                    //データリーダーを取得する
                    dR = dCon.FreeReader(sqlSTRING);

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    while (dR.Read())
                    {
                        cmb1 = new ComboShain();
                        cmb1.ID = int.Parse(dR["shainid"].ToString());
                        cmb1.DisplayName = dR["shaincode"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
                        cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
                        cmb1.code = (dR["shaincode"].ToString() + "").Trim();
                        cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
                        cmb1.BumonCode = dR["bumoncode"].ToString().PadLeft(3, '0');
                        cmb1.BumonName = dR["bumonname"].ToString();
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    dCon.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "社員コンボボックスロード");
                }

            }
        }


        // データ領域コンボボックスクラス
        public class ComboDataArea
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            // データ領域ロード
            public static void load(ComboBox tempObj)
            {
                dbControl.DataControl dcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
                OleDbDataReader dR = null;

                try
                {
                    ComboDataArea cmb;

                    // データリーダー取得
                    string mySql = string.Empty;
                    mySql += "SELECT * FROM Common_Unit_DataAreaInfo ";
                    mySql += "where CompanyTerm = " + DateTime.Today.Year.ToString();
                    dR = dcon.FreeReader(mySql);

                    //会社情報がないとき
                    if (!dR.HasRows)
                    {
                        MessageBox.Show("会社領域情報が存在しません", "会社領域選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    // コンボボックスにアイテムを追加します
                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";

                    while (dR.Read())
                    {
                        cmb = new ComboDataArea();
                        // "CompanyCode"が数字のレコードを対象とする
                        if (Utility.NumericCheck(dR["CompanyCode"].ToString()))
                        {
                            cmb.DisplayName = dR["CompanyName"].ToString().Trim();
                            cmb.ID = dR["Name"].ToString().Trim();
                            cmb.code = dR["CompanyCode"].ToString().Trim();
                            tempObj.Items.Add(cmb);
                        }
                    }
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

            }
        }


        ///--------------------------------------------------------
        /// <summary>
        /// 会社情報より部門コード桁数、社員コード桁数を取得
        /// </summary>
        /// -------------------------------------------------------
        public class BumonShainKetasu
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            // 会社情報取得
            public static void GetKetasu(string dbName)
            {
                dbControl.DataControl dcon = new dbControl.DataControl(dbName);
                OleDbDataReader dR = null;

                try
                {
                    // データリーダー取得
                    string mySql = string.Empty;
                    mySql += "SELECT BumonCodeKeta,ShainCodeKeta FROM Kaisha ";
                    dR = dcon.FreeReader(mySql);

                    // 部門コード桁数、社員コード桁数を取得
                    while (dR.Read())
                    {
                        global.ShozokuLength = int.Parse(dR["BumonCodeKeta"].ToString());
                        global.ShainLength = int.Parse(dR["ShainCodeKeta"].ToString());
                    }
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

            }
        }


        ///------------------------------------------------------------------
        /// <summary>
        ///     ファイル選択ダイアログボックスの表示 </summary>
        /// <param name="sTitle">
        ///     タイトル文字列</param>
        /// <param name="sFilter">
        ///     ファイルのフィルター</param>
        /// <returns>
        ///     選択したファイル名</returns>
        ///------------------------------------------------------------------
        public static string userFileSelect(string sTitle, string sFilter)
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = sTitle;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = sFilter;
            //openFileDialog1.Filter = "CSVファイル(*.CSV)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        public class frmMode
        {
            public int ID { get; set; }

            public int Mode { get; set; }

            public int rowIndex { get; set; }
        }

        public class xlsShain
        {
            public int sCode { get; set; }
            public string sName { get; set; }
            public int bCode { get; set; }
            public string bName { get; set; }
            public int kbn { get; set; }
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを追加モードで出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            //// ファイル名（タイムスタンプ）
            //string timeStamp = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            //// ファイル名
            //string outFileName = sPath + timeStamp + ".csv";

            //// 出力ファイルが存在するとき
            //if (System.IO.File.Exists(outFileName))
            //{
            //    // 既存のファイルを削除
            //    System.IO.File.Delete(outFileName);
            //}

            // CSVファイル出力
            //System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("shift-jis"));
            System.IO.File.AppendAllLines(sPath, arrayData, Encoding.GetEncoding("shift-Jis"));
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string getStringSubMax(string s, int n)
        {
            string val = string.Empty;

            // 文字間のスペースを除去 2015/03/10
            s = s.Replace(" ", "");

            if (s.Length > n)
            {
                val = s.Substring(0, n);
            }
            else
            {
                val = s;
            }

            return val;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     暦補正値を取得する：2019/04/25 </summary>
        /// <returns>
        ///     RekiHosei</returns>
        ///----------------------------------------------------------
        public static int GetRekiHosei() => Properties.Settings.Default.RekiHosei;


        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

    }
}
