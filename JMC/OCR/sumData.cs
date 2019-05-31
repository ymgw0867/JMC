using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.IO;
using JMC.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace JMC.OCR
{
    class sumData
    {
        public sumData()
        {
            for (int i = 0; i < c41.Length; i++)
            {
                if (i != c41.Length - 1) c41[i] = ",";
            }
        }

        public void Caltotal(OleDbDataReader dR)
        {
            cCnt++;
            cYear = dR["年"].ToString();
            cMonth = dR["月"].ToString();
            KojinNum = dR["個人番号"].ToString();
            ShainID = dR["社員ID"].ToString();
            KyuyoKbn = dR["給与区分"].ToString();

            // パートタイマーは労働時間を計算する
            if (dR["給与区分"].ToString() == "1")
            {
                SouroudouH += Utility.StrtoInt(dR["総労働"].ToString());
                SouroudouM += Utility.StrtoInt(dR["総労働分"].ToString());
            }
            else
            {
                // 社員は労働時間は計算しない
                SouroudouH = 0;
                SouroudouM = 0;
            }

            // パートタイマー以外は欠勤日数を計算する
            if (dR["給与区分"].ToString() == "1")
                Kekkin = 0;
            else Kekkin += Utility.StrtoInt(dR["欠勤日数合計"].ToString());

            TokkyuN += Utility.StrtoInt(dR["特休日数合計"].ToString());
            FurikyuN += Utility.StrtoInt(dR["振休日数合計"].ToString());
            ZangyoH += Utility.StrtoInt(dR["残業時"].ToString());
            ZangyoM += Utility.StrtoInt(dR["残業分"].ToString());
            ShinyaTime += Utility.StrtoInt(dR["深夜勤務時間合計"].ToString());
            Chisou += Utility.StrtoInt(dR["遅刻早退回数"].ToString());
            YukyuN += Utility.StrtoInt(dR["有休日数合計"].ToString());
            YukyuTime += Utility.StrtoInt(dR["有休時間合計"].ToString());

            kitei = Utility.StrtoInt(dR["月間規定勤務時間"].ToString());    // 2014/10/28
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     常陽コンピュータサービス向け給与データ集計 </summary>
        /// <param name="dR">
        ///     出勤簿データリーダー</param>
        ///     
        ///     2016.11.18
        ///     
        /// 2019/04/25 : 年西暦化    
        ///----------------------------------------------------------------------
        public void CaltotalJcs(OleDbDataReader dR)
        {
            jCnt++;
            //jYear = dR["年"].ToString();   // 2019/04/25 コメント化
            jYear = (Utility.StrtoInt(dR["年"].ToString()) + Utility.GetRekiHosei()).ToString();  // 2019/04/25 西暦年に変更
            jMonth = dR["月"].ToString();

            // メインの勤務先コードを採用する
            if (dR["勤務先区分"].ToString() == global.KINMU_MAIN.ToString())
            {
                jZCode = dR["所属コード"].ToString();
            }

            jShainNum = dR["個人番号"].ToString();

            // パートタイマー
            if (dR["給与区分"].ToString() == global.STATUS_PART.ToString())
            {
                // 労働時間を計算する
                if (jCnt == 1)
                {
                    jSouroudouTM1 += Utility.StrtoInt(dR["総労働"].ToString()) * 60 + Utility.StrtoInt(dR["総労働分"].ToString());
                }
                else if (jCnt == 2)
                {
                    jSouroudouTM2 += Utility.StrtoInt(dR["総労働"].ToString()) * 60 + Utility.StrtoInt(dR["総労働分"].ToString());
                }
                else if (jCnt == 3)
                {
                    jSouroudouTM3 += Utility.StrtoInt(dR["総労働"].ToString()) * 60 + Utility.StrtoInt(dR["総労働分"].ToString());
                }

                // 欠勤日数なし
                jKekkin = 0;
            }
            else
            {
                // 社員は労働時間は計算しない
                jSouroudouTM1 = 0;
                jSouroudouTM2 = 0;
                jSouroudouTM3 = 0;

                // 社員は欠勤日数を計算する
                jKekkin += Utility.StrtoInt(dR["欠勤日数合計"].ToString());
            }

            // メイン出勤簿のとき
            if (Utility.StrtoInt(dR["勤務先区分"].ToString()) == global.KINMU_MAIN)
            {
                // 出勤日数
                jShukkinN = getShukkinNisu(dR["個人番号"].ToString(), Utility.StrtoInt(dR["給与区分"].ToString()), global.KINMU_MAIN);

                // 総出勤日数
                jShukkinNTL = getShukkinNisu(dR["個人番号"].ToString(), Utility.StrtoInt(dR["給与区分"].ToString()), global.flgOff);

                // 有給休暇日数
                jYukyuN = dR["有休日数合計"].ToString();

                // 残業時間・分単位
                jZangyoH = Utility.StrtoInt(dR["残業時"].ToString()) * 60 + Utility.StrtoInt(dR["残業分"].ToString());

                //// 深夜勤務時間・分単位
                //jShinyaTime = Utility.StrtoInt(dR["深夜勤務時間合計"].ToString());

                // 立替金
                jTatekae = Utility.StrtoInt(dR["立替金"].ToString());

                // 旅費交通費
                jRyohi = Utility.StrtoInt(dR["旅費交通費"].ToString());

                // その他支給
                jsonota = Utility.StrtoInt(dR["その他支給"].ToString());

                // 特休日数
                jTokkyuN = Utility.StrtoInt(getNisu(dR["個人番号"].ToString(), Utility.StrtoInt(dR["給与区分"].ToString()), global.flgOff, global.TOKUBETSU_KYUKA));

                // 振出日数
                jFurideN = Utility.StrtoInt(getNisu(dR["個人番号"].ToString(), Utility.StrtoInt(dR["給与区分"].ToString()), global.flgOff, global.FURIDE_KYUKA));

                // 振休日数
                jFurikyuN = Utility.StrtoInt(getNisu(dR["個人番号"].ToString(), Utility.StrtoInt(dR["給与区分"].ToString()), global.flgOff, global.FURIKYU_KYUKA));
            }

            // 深夜勤務時間・分単位 2017/06/21  メイン・サブ両方対象とする
            jShinyaTime += Utility.StrtoInt(dR["深夜勤務時間合計"].ToString());

            //jTokkyuN += Utility.StrtoInt(dR["特休日数合計"].ToString());
            //jFurideN += Utility.StrtoInt(dR["振出日数合計"].ToString());
            //jFurikyuN += Utility.StrtoInt(dR["振休日数合計"].ToString());
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     受渡データ出力 </summary>
        /// <param name="outFile">
        ///     出力するStreamWriterオブジェクト</param>
        /// <param name="sd">
        ///     集計データクラス</param>
        /// ------------------------------------------------------------------------
        public void SaveDatacsv(StreamWriter outFile)
        {
            c1 = KojinNum + ",";

            // 出勤日数
            string sN = getShukkinNisu(ShainID, int.Parse(KyuyoKbn), global.flgOff);
            c4 = sN + ",";

            // 総労働時間
            double sh = 0; 
            int sm = 0;
            if (SouroudouM >= 60)
            {
                SouroudouH += (int)System.Math.Floor(double.Parse(SouroudouM.ToString()) / 60);
                SouroudouM = int.Parse(SouroudouM.ToString()) % 60;
            }
            c5 = SouroudouH.ToString() + ":" + Utility.StrtoInt(SouroudouM.ToString()).ToString().PadLeft(2, '0') + ",";
            
            c7 = Kekkin.ToString() + ",";    // 欠勤日数
            c8 = (Utility.StrtoInt(TokkyuN.ToString()) + FurikyuN).ToString() + ",";    // 特休日数

            // 残業時間
            if (ZangyoM >= 60)
            {
                ZangyoH += (int)System.Math.Floor(double.Parse(ZangyoM.ToString()) / 60);
                ZangyoM = int.Parse(ZangyoM.ToString()) % 60;
            }

            c12 = ZangyoH.ToString() + ":" + ZangyoM.ToString().PadLeft(2, '0') + ",";

            // 深夜勤務時間
            sh = System.Math.Floor(double.Parse(ShinyaTime.ToString()) / 60);
            sm = int.Parse(ShinyaTime.ToString()) % 60;
            c13 = sh.ToString() + ":" + sm.ToString().PadLeft(2, '0') + ",";

            c18 = Chisou.ToString() + ",";          // 遅刻早退
            c20 = YukyuN.ToString() + ",";          // 有給日数
            c21 = YukyuTime.ToString() + ",";       // 有給時間

            // 要勤務日数
            c2 = (int.Parse(sN) + Utility.StrtoInt(YukyuN.ToString())).ToString() + ",";

            // CSVファイルを書き出す
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append(c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9 + c10);
            sb.Append(c11 + c12 + c13 + c14 + c15 + c16 + c17 + c18 + c19 + c20);
            sb.Append(c21 + c22 + c23 + c24 + c25 + c26 + c27 + c28 + c29 + c30);
            sb.Append(c31 + c32 + c33 + c34 + c35 + c36 + c37 + c38 + c39 + c40);

            for (int i = 0; i < c41.Length; i++)
            {
                sb.Append(c41[i]);
            }

            //明細ファイル出力
            outFile.WriteLine(sb.ToString());
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     常陽コンピュータサービス向け受渡データ出力 </summary>
        /// <param name="xls">
        ///     出力配列 </param>
        /// <param name="i">
        ///     配列インデックス</param>
        /// ------------------------------------------------------------------------
        public void SaveDataJcs(ref string [] xls, int i, Utility.xlsShain [] shainArray)
        {
            string sCode = string.Empty;
            string sName = string.Empty;

            // CSVファイルを書き出す
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append(jYear).Append(",");
            sb.Append(jMonth).Append(",");
            sb.Append(jZCode).Append(",");
            sb.Append(jShainNum).Append(",");

            // 社員名取得
            sb.Append(Utility.ComboShain.getXlsSname(shainArray, Utility.StrtoInt(jShainNum),  out sCode, out sName)).Append(",");

            //sb.Append(jShukkinN).Append(",");     // 2017/06/21 出勤日数撤廃
            sb.Append(jShukkinNTL).Append(",");
            sb.Append(jSouroudouTM1).Append(",");
            sb.Append(jSouroudouTM2).Append(",");
            sb.Append(jSouroudouTM3).Append(",");
            sb.Append(jYukyuN).Append(",");
            sb.Append(jKekkin).Append(",");
            sb.Append(jTokkyuN).Append(",");
            sb.Append(jFurideN).Append(",");
            sb.Append(jFurikyuN).Append(",");
            sb.Append(jZangyoH).Append(",");
            //sb.Append(jShinyaTime).Append(",");     // 2017/06/21 深夜残業時間撤廃
            sb.Append(jShinyaTime);     // 2017/05/07

            //// 以下 2017/05/07
            //sb.Append(jTatekae).Append(",");
            //sb.Append(jRyohi).Append(",");
            //sb.Append(jsonota);

            // 配列に格納
            Array.Resize(ref xls, i + 1);
            xls[i] = sb.ToString();
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     社員別集計結果配列を多次元配列に展開する </summary>
        /// <param name="mAry">
        ///     多次元配列</param>
        /// <param name="sAry">
        ///     集計結果配列</param>
        ///-------------------------------------------------------------
        public void arrayStoMl(ref string [,] mAry, string [] sAry)
        {
            int r = sAry.Length;
            int c = sAry[0].Split(',').Count();

            mAry = new string[r, c];

            int rX = 0;
            foreach (var t in sAry)
            {
                string [] cc = t.Split(',');

                mAry[rX, 0] = cc[0];
                mAry[rX, 1] = cc[1];
                mAry[rX, 2] = cc[2];
                mAry[rX, 3] = cc[3];
                mAry[rX, 4] = cc[4];
                mAry[rX, 5] = cc[5];
                mAry[rX, 6] = cc[6];
                mAry[rX, 7] = cc[7];
                mAry[rX, 8] = cc[8];
                mAry[rX, 9] = cc[9];
                mAry[rX, 10] = cc[10];
                mAry[rX, 11] = cc[11];
                mAry[rX, 12] = cc[12];
                mAry[rX, 13] = cc[13];
                mAry[rX, 14] = cc[14];
                mAry[rX, 15] = cc[15];
                mAry[rX, 16] = cc[16];
                mAry[rX, 17] = cc[17];
                mAry[rX, 18] = cc[18];
                mAry[rX, 19] = cc[19];

                rX++;
            }
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     社員別集計ＣＳＶデータを多次元配列に展開する </summary>
        /// <param name="mAry">
        ///     多次元配列</param>
        /// <param name="sFile">
        ///     ＣＳＶファイルパス</param>
        ///-------------------------------------------------------------
        public void csvToMlArray(ref string[,] mAry, string sFile)
        {
            string [] sArray = System.IO.File.ReadAllLines(sFile, Encoding.Default);
            
            int r = sArray.Length;
            int c = sArray[0].Split(',').Count();

            mAry = new string[r, c];

            int rX = 0;
            foreach (var t in sArray)
            {
                string[] cc = t.Split(',');

                mAry[rX, 0] = cc[0];
                mAry[rX, 1] = cc[1];
                mAry[rX, 2] = cc[2];
                mAry[rX, 3] = cc[3];
                mAry[rX, 4] = cc[4];
                mAry[rX, 5] = cc[5];
                mAry[rX, 6] = cc[6];
                mAry[rX, 7] = cc[7];
                mAry[rX, 8] = cc[8];
                mAry[rX, 9] = cc[9];
                mAry[rX, 10] = cc[10];
                mAry[rX, 11] = cc[11];
                mAry[rX, 12] = cc[12];
                mAry[rX, 13] = cc[13];
                mAry[rX, 14] = cc[14];
                mAry[rX, 15] = cc[15];

                //mAry[rX, 16] = cc[16];    // 2017/06/21
                //mAry[rX, 17] = cc[17];    // 2017/06/21

                //mAry[rX, 18] = cc[18];    // 2017/05/08
                //mAry[rX, 19] = cc[19];    // 2017/05/08

                rX++;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     常陽コンピュータサービス向けエクセルシート作成 </summary>
        /// <param name="xls">
        ///     エクセルシート</param>
        /// <param name="xlsFile">
        ///     常陽コンピュータサービス給与エクセルシート</param>
        ///------------------------------------------------------------------
        public void saveExcelKyuyo(string [,] xls, string xlsFile)
        {
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            const int sGYO = 3;       //エクセルファイル明細開始行

            try
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;
                oXls.DisplayAlerts = false;

                ////// 前回の書き込みセルを初期化する
                ////rng = oxlsSheet.Range[oxlsSheet.Cells[sGYO, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                ////rng.Value2 = "";

                //// 給与データ書き込み
                //rng = oxlsSheet.Range[oxlsSheet.Cells[sGYO, 1], oxlsSheet.Cells[(sGYO + xls.GetLength(0) - 1), oxlsSheet.UsedRange.Columns.Count]];
                //rng.Value2 = xls;
                
                // 個別セル設定：2017/06/21
                for (int i = 0; i < xls.GetLength(0); i++)
                {
                    for (int x = 0; x < xls.GetLength(1); x++)
                    {
                        oxlsSheet.Cells[sGYO + i, x + 1].Value = xls[i, x];
                    }
                }
                
                // シートを保存
                oXlsBook.SaveAs(xlsFile, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);                
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "給与データエクセル出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
            }
            finally
            {
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
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     出勤日数取得 </summary>
        /// <param name="sID">
        ///     社員ID</param>
        /// <param name="Yaku">
        ///     役職タイプ</param>
        /// <param name="main">
        ///     1:メイン出勤簿のみ対象, 0:全ての出勤簿を対象</param> // 2016/11/18
        /// <returns>
        ///     出勤日数</returns>
        ///----------------------------------------------------------------
        private string getShukkinNisu(string sID, int Yaku, int main)
        {
            // 出力データ生成
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            // 日付クラスインスタンス生成
            dataByDay[] dby = new dataByDay[31];
            for (int i = 0; i < dby.Length; i++)
            {
                dby[i] = new dataByDay();
                dby[i].day = 0;
                dby[i].yukyu = string.Empty;
                dby[i].sH = string.Empty;
                dby[i].sM = string.Empty;
                dby[i].eH = string.Empty;
                dby[i].eM = string.Empty;
            }

            // 出勤簿明細データリーダーを取得します
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT 出勤簿明細.* from 出勤簿ヘッダ inner join 出勤簿明細 ");
            sb.Append("on 出勤簿ヘッダ.ID = 出勤簿明細.ヘッダID ");
            //sb.Append("where 出勤簿ヘッダ.社員ID = " + sID + " ");
            sb.Append("where 出勤簿ヘッダ.個人番号 = '" + sID + "' ");   // 2016/11/17

            // メイン出勤簿のみを対象とする場合 2016/11/18
            if (main == global.KINMU_MAIN)
            {
                sb.Append("and 勤務先区分 = " + global.KINMU_MAIN.ToString() + " ");
            }

            sb.Append("order by 出勤簿明細.ID ");
            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            // 勤務記録を日付配列にセットします
            while (dR.Read())
            {
                // 日付
                int iX = int.Parse(dR["日付"].ToString()) - 1;
                dby[iX].day = int.Parse(dR["日付"].ToString());

                // 有給休暇
                string yu = Utility.NulltoStr(dR["有給記号"].ToString());
                if (yu != string.Empty)
                {
                    dby[iX].yukyu = yu;
                }

                // 開始・終了時刻
                if (Utility.NulltoStr(dR["開始時"].ToString()) != string.Empty &&
                    Utility.NulltoStr(dR["開始時"].ToString()) != "24")
                {
                    dby[iX].sH = Utility.NulltoStr(dR["開始時"].ToString());
                    dby[iX].sM = Utility.NulltoStr(dR["開始分"].ToString());
                    dby[iX].eH = Utility.NulltoStr(dR["終了時"].ToString());
                    dby[iX].eM = Utility.NulltoStr(dR["終了分"].ToString());
                }
            }

            dR.Close();
            sCom.Connection.Close();

            // 出勤日数初期化
            int sDays = 0;

            // 日付配列を読む
            for (int i = 0; i < dby.Length; i++)
            {
                if (dby[i].day != 0)
                {
                    // 勤務時間が記入されている日
                    if (dby[i].sH != string.Empty && dby[i].sM != string.Empty)
                    {
                        // 開始時間が24時台以外のもの（24時台は前日からの通し勤務とみなし出勤日数に加えない）
                        if (dby[i].sH != "24")
                        {
                            if (Yaku == global.STATUS_SHAIN)
                            {
                                sDays++;  // 社員
                            }
                            else if (Utility.NulltoStr(dby[i].yukyu) != global.ZENNICHI_YUKYU)
                            {
                                // パート：終日有休以外のときは出勤日数としてカウントする
                                sDays++;
                            }
                        }
                    }
                }
            }

            return sDays.ToString();
        }



        ///----------------------------------------------------------------
        /// <summary>
        ///     同日を含まない休暇日数取得 </summary>
        /// <param name="sID">
        ///     社員ID</param>
        /// <param name="Yaku">
        ///     役職タイプ</param>
        /// <param name="main">
        ///     1:メイン出勤簿のみ対象, 0:全ての出勤簿を対象</param> // 2016/11/18
        /// <param name="kigou">
        ///     日数を取得する休暇記号</param>
        /// <returns>
        ///     出勤日数</returns>
        ///----------------------------------------------------------------
        private string getNisu(string sID, int Yaku, int main, string kigou)
        {
            // 出力データ生成
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            // 日付クラスインスタンス生成
            dataByDay[] dby = new dataByDay[31];
            for (int i = 0; i < dby.Length; i++)
            {
                dby[i] = new dataByDay();
                dby[i].day = 0;
                dby[i].yukyu = string.Empty;
                dby[i].sH = string.Empty;
                dby[i].sM = string.Empty;
                dby[i].eH = string.Empty;
                dby[i].eM = string.Empty;
            }

            // 出勤簿明細データリーダーを取得します
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT 出勤簿明細.* from 出勤簿ヘッダ inner join 出勤簿明細 ");
            sb.Append("on 出勤簿ヘッダ.ID = 出勤簿明細.ヘッダID ");
            sb.Append("where 出勤簿ヘッダ.個人番号 = '" + sID + "' ");   // 2016/11/17

            // メイン出勤簿のみを対象とする場合 2016/11/18
            if (main == global.KINMU_MAIN)
            {
                sb.Append("and 勤務先区分 = " + global.KINMU_MAIN.ToString() + " ");
            }

            sb.Append("order by 出勤簿明細.ID ");
            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            // 勤務記録を日付配列にセットします
            while (dR.Read())
            {
                // 日付
                int iX = int.Parse(dR["日付"].ToString()) - 1;
                dby[iX].day = int.Parse(dR["日付"].ToString());

                // 指定の休暇記号に該当するか
                if (Utility.NulltoStr(dR["休暇記号"].ToString()) == kigou)
                {
                    dby[iX].sH = global.flgOn.ToString();
                }
            }

            dR.Close();
            sCom.Connection.Close();

            // 出勤日数初期化
            int sDays = 0;

            // 日付配列を読む
            for (int i = 0; i < dby.Length; i++)
            {
                if (dby[i].day != 0)
                {
                    // 指定記号該当日のとき
                    if (dby[i].sH != string.Empty)
                    {
                        sDays++;
                    }
                }
            }

            return sDays.ToString();
        }



        // 集計エリア
        public int cCnt = 0;                    // 勤務票枚数 2014/10/28
        public string cYear = string.Empty;     // 2014/10/28
        public string cMonth = string.Empty;    // 2014/10/28
        public string KojinNum = string.Empty;
        public string ShainID = string.Empty;
        public string KyuyoKbn = string.Empty;
        public int SouroudouH = 0;
        public int SouroudouM = 0;
        public int TokkyuN = 0;
        public int FurikyuN = 0;
        public int ZangyoH = 0;
        public int ZangyoM = 0;
        public int ShinyaTime = 0;
        public int Chisou = 0;
        public int Kekkin = 0;
        public int YukyuN = 0;
        public int YukyuTime = 0;
        public int kitei = 0;   // 2014/10/28

        // ＰＣＡ給与分散データ（勤怠汎用データ）レイアウト
        public string c1 = string.Empty;    // 社員コード
        public string c2 = string.Empty;    // 要勤務日数
        public string c3 = ",";             // 要勤務時間
        public string c4 = string.Empty;    // 出勤日数
        public string c5 = string.Empty;    // 出勤時間
        public string c6 = "0,";            // 事故欠勤日数
        public string c7 = string.Empty;    // 病気欠勤日数
        public string c8 = string.Empty;    // 代休特休時間
        public string c9 = ",";             // 休日出勤日数
        public string c10 = ",";            // 有休消化日数
        public string c11 = ",";            // 有休残日数
        public string c12 = string.Empty;   // 残業平日普通
        public string c13 = string.Empty;   // 残業平日深夜
        public string c14 = ",";            // 残業休日普通
        public string c15 = ",";            // 残業休日深夜
        public string c16 = ",";            // 残業法定普通
        public string c17 = ",";            // 残業法定深夜
        public string c18 = string.Empty;   // 遅刻早退回数
        public string c19 = ",";            // 遅刻早退時間
        public string c20 = string.Empty;   // 有休日数消化
        public string c21 = string.Empty;   // 有休時間消化
        public string c22 = ",";            // 有休日数残
        public string c23 = ",";            // 有休時間残
        public string c24 = ",";            // 有休可能時間
        public string c25 = ",";            // 残業平日普通４５下
        public string c26 = ",";            // 残業平日普通４５超
        public string c27 = ",";            // 残業平日普通６０超
        public string c28 = ",";            // 残業平日普通代休
        public string c29 = ",";            // 残業平日深夜４５下
        public string c30 = ",";            // 残業平日深夜４５超
        public string c31 = ",";            // 残業平日深夜６０超
        public string c32 = ",";            // 残業平日深夜代休
        public string c33 = ",";            // 残業平日普通４５下
        public string c34 = ",";            // 残業平日普通４５超
        public string c35 = ",";            // 残業平日普通６０超
        public string c36 = ",";            // 残業平日普通代休
        public string c37 = ",";            // 残業平日深夜４５下
        public string c38 = ",";            // 残業平日深夜４５超
        public string c39 = ",";            // 残業平日深夜６０超
        public string c40 = ",";            // 残業平日深夜代休

        public string [] c41 = new string [50];

        // 常陽コンピュータサービス向け出力データ
        public int jCnt = 0;                    // 勤務票枚数
        public string jYear = string.Empty;     // 2014/10/28
        public string jMonth = string.Empty;    // 2014/10/28
        public string jZCode = string.Empty;    // 所属コード
        public string jShainNum = string.Empty; // 社員番号
        public string jShukkinN = "0";      // 出勤日数
        public string jShukkinNTL = "0";    // 総出勤日数
        public int jSouroudouTM1 = 0;   // 勤務時間１
        public int jSouroudouTM2 = 0;   // 勤務時間２
        public int jSouroudouTM3 = 0;   // 勤務時間３
        public string jYukyuN = "0";    // 有給休暇日数
        public int jKekkin = 0;         // 欠勤日数合計
        public int jTokkyuN = 0;        // 特休日数合計
        public int jFurideN = 0;        // 振替出勤日数合計
        public int jFurikyuN = 0;       // 振替休日日数合計
        public int jZangyoH = 0;        // 残業時間
        public int jShinyaTime = 0;     // 深夜勤務時間
        public int jTatekae = 0;        // 立替金
        public int jRyohi = 0;          // 旅費交通費
        public int jsonota = 0;         // その他支給


    }
}
