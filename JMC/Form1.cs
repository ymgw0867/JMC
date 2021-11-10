using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JMC.Common;
using JMC.Config;
using JMC.OCR;

namespace JMC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // キャプションにバージョンを追加 : 2021/11/10
            this.Text += "  ver " + Application.ProductVersion;

            // 表示サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);  // 最大サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);  // 最小サイズ

            // 過去出勤簿ヘッダ＠所属コード桁数４ケタ　2013/06/13
            // mdbSzkLenAlter(); // コメント化：2021/11/10
        }

        /// <summary>
        /// 
        /// 2013/06/13
        /// ローカルＭＤＢ 
        /// 過去出勤簿ヘッダ＠所属コード桁数４ケタ化
        /// 
        /// </summary>
        private void mdbSzkLenAlter()
        {
            // ローカルデータベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;
            sCom.Connection = Con.cnOpen();

            DataTable dt = null;
            int hLen = 0;
            int sLen = 0;

            // 過去出勤簿ヘッダの所属コード桁数を調べる
            string sqlSTRING = string.Empty;
            sqlSTRING += "select * from 過去出勤簿ヘッダ";
            sCom.CommandText = sqlSTRING;
            dR = sCom.ExecuteReader();
            dt = dR.GetSchemaTable();
            foreach (DataRow rw in dt.Rows)
            {
                if (rw[0].ToString() == "所属コード")
                {
                    hLen = int.Parse(rw[2].ToString());
                    break;
                }
            }
            dR.Close();
            dt.Dispose();

            // 出勤簿ヘッダの所属コード桁数を調べる
            sqlSTRING = string.Empty;
            sqlSTRING += "select * from 出勤簿ヘッダ";
            sCom.CommandText = sqlSTRING;
            dR = sCom.ExecuteReader();
            dt = dR.GetSchemaTable();
            foreach (DataRow rw in dt.Rows)
            {
                if (rw[0].ToString() == "所属コード")
                {
                    sLen = int.Parse(rw[2].ToString());
                    break;
                }
            }
            dR.Close();
            dt.Dispose();

            // 過去出勤簿ヘッダの所属コードが３桁なら桁数を10桁に変更する
            if (hLen == 3)
            {
                sqlSTRING = string.Empty;
                sqlSTRING += "ALTER TABLE 過去出勤簿ヘッダ ALTER COLUMN 所属コード TEXT(10)";
                sCom.CommandText = sqlSTRING;
                sCom.ExecuteNonQuery();
            }
            
            // 出勤簿ヘッダの所属コードが３桁なら桁数を10桁に変更する
            if (sLen == 3)
            {
                sqlSTRING = string.Empty;
                sqlSTRING += "ALTER TABLE 出勤簿ヘッダ ALTER COLUMN 所属コード TEXT(10)";
                sCom.CommandText = sqlSTRING;
                sCom.ExecuteNonQuery();
            }

            sCom.Connection.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmCalendar frm = new frmCalendar();
            frm.ShowDialog();
            this.Show();
        }
       
        
        private void button1_Click(object sender, EventArgs e)
        {
            // プレ印刷
            this.Hide();
            frmComSelect frm = new frmComSelect(0);
            frm.ShowDialog();

            if (frm._pFileName != string.Empty)
            {
                // 選択領域のファイル名を取得します
                string _grpID = frm._pID;
                string _grpName = frm._PName;
                string _exfileName = frm._pFileName;
                string _exSheetNum = frm._pSheetNum;
                int _yakushokuType = frm._pYakushokuType;
                
                frm.Dispose();

                // 全てＣＳＶファイルで管理 2017/05/17
                int pMode = global.CSV_MODE;
                
                //int pMode = 0;

                //if (_grpID == "")
                //{
                //    // CSVモード
                //    pMode = global.CSV_MODE;
                //}
                //else
                //{
                //    // エクセルシートモード
                //    pMode = global.XLS_MODE;
                //}


                // 出勤簿プレ印刷
                prePrint.prePrint frmP = new prePrint.prePrint(_exfileName, _exSheetNum, _yakushokuType, _grpID, _grpName, pMode);
                frmP.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect(1);
            frm.ShowDialog();

            if (frm._pFileName != string.Empty)
            {
                // 選択領域のファイル名を取得します
                string _exfileName = frm._pFileName;
                string _exSheetNum = frm._pSheetNum;
                int _yakushokuType = frm._pYakushokuType;
                string _grpID = frm._pID;
                
                frm.Dispose();

                // 月給者月間勤務時間設定
                Config.frmGekkyuKinmu frmg = new frmGekkyuKinmu(_grpID);
                frmg.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmOcr frm = new frmOcr();
            frm.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect(1);
            frm.ShowDialog();

            if (frm._pFileName != string.Empty)
            {
                // 選択領域のファイル名を取得します
                string _grpID = frm._pID;
                string _exfileName = frm._pFileName;
                string _exSheetNum = frm._pSheetNum;
                int _yakushokuType = frm._pYakushokuType;

                frm.Dispose();

                // 出勤簿データ作成画面
                OCR.frmCorrect frmg = new frmCorrect(_exfileName, _exSheetNum, _yakushokuType, _grpID, string.Empty);
                frmg.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmUnSubmit frm = new frmUnSubmit();
            frm.ShowDialog();
            this.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            JMC.Config.frmShainFile frm = new frmShainFile();
            frm.ShowDialog();
            this.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmPattern frm = new frmPattern();
            frm.ShowDialog();
            this.Show();
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("常陽コンピュータサービス向けエクセル給与シート出力を行いますか？","確認",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            this.Hide();
            frmComSelect frm = new frmComSelect(1);
            frm.ShowDialog();

            if (frm._pFileName != string.Empty)
            {
                // 選択領域のファイル名を取得します
                string _grpID = frm._pID;
                string _gName = frm._PName;
                string _exfileName = frm._pFileName;
                string _exSheetNum = frm._pSheetNum;
                int _yakushokuType = frm._pYakushokuType;

                frm.Dispose();

                // 常陽コンピュータサービス向けエクセル給与シート出力 
                string sPath = Properties.Settings.Default.instPath + _grpID.PadLeft(3, '0') + " " + _gName + @"\";
                
                // 出勤簿集計ＣＳＶデータパス
                if (System.IO.File.Exists(sPath + Properties.Settings.Default.outCsvName))
                {
                    setExcelData(sPath);
                    MessageBox.Show("常陽コンピュータサービス向けエクセル給与シート出力が終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("出勤簿集計ＣＳＶデータが存在しません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                frm.Dispose();
            }

            this.Show();
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     常陽コンピュータサービス向けエクセル給与シート出力を実行する </summary>
        /// <param name="gPath">
        ///     グループフォルダパス</param>
        ///---------------------------------------------------------------------
    private void setExcelData(string gPath)
        {
            string csvPath = gPath + Properties.Settings.Default.outCsvName;  // 出勤簿集計ＣＳＶデータパス
            string gXlsPath = gPath + Properties.Settings.Default.outXlsJcs;  // エクセル給与シートパス

            sumData sd = new sumData();

            // 出勤簿集計ＣＳＶデータを多次元配列へセット
            string[,] mArray = null;
            sd.csvToMlArray(ref mArray, csvPath);

            //// エクセル給与シートが存在したら削除する  2017/05/08
            //if (System.IO.File.Exists(gXlsPath))
            //{
            //    System.IO.File.Delete(gXlsPath);
            //}

            // エクセル給与シートが存在したら名前を変更して保存する  2017/05/08
            DateTime dn = DateTime.Now;
            string fdt = string.Empty;
            string newFnm = string.Empty;

            fdt = dn.Year.ToString() + dn.Month.ToString().PadLeft(2, '0') +
                dn.Day.ToString().PadLeft(2, '0') + dn.Hour.ToString().PadLeft(2, '0') +
                dn.Minute.ToString().PadLeft(2, '0') + dn.Second.ToString().PadLeft(2, '0');

            if (System.IO.File.Exists(gXlsPath))
            {
                newFnm = System.IO.Path.GetFileNameWithoutExtension(gXlsPath) + fdt + ".xlsx";                
                System.IO.File.Move(gXlsPath, gPath + newFnm);
            }

            // エクセル給与シートをグループフォルダへコピーする
            //MessageBox.Show(Properties.Settings.Default.xlsPath + Properties.Settings.Default.outXlsJcs + " " + gXlsPath);
            System.IO.File.Copy(Properties.Settings.Default.xlsPath + Properties.Settings.Default.outXlsJcs, gXlsPath);
          
            // 常陽コンピュータサービス向けエクセル給与シート出力 2016/12/06
            sd.saveExcelKyuyo(mArray, gXlsPath);

            //// 出勤簿集計ＣＳＶデータを削除する
            //System.IO.File.Delete(csvPath);

            // 出勤簿集計ＣＳＶデータを名前を変更して保存する
            newFnm = System.IO.Path.GetFileNameWithoutExtension(csvPath) + fdt + ".csv";
            System.IO.File.Move(csvPath, gPath + newFnm);
        }
    }
}
