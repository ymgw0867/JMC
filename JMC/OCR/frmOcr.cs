using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Imaging;
using IdrFormEngine;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using JMC.Common;

namespace JMC.OCR
{
    public partial class frmOcr : Form
    {
        // OCR変換枚数
        int _pageCount;

        public frmOcr()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif").Count() == 0)
            {
                MessageBox.Show("OCR認識対象となる画像がありません", "画像なし", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("スキャナで作成された出勤簿画像データのOCR変換処理を実施します。よろしいですか？","実行確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // ＯＣＲ認識処理
            GetOcrText();


            //// マルチTIFを分解する
            //MultiTif(Properties.Settings.Default.scanPath);

            //// OCR認識処理を実施する
            //ocrMain(Properties.Settings.Default.readPath,
            //        Properties.Settings.Default.ngPath,
            //        Properties.Settings.Default.dataPath,
            //        Properties.Settings.Default.fmtFilePath,
            //        _pageCount);

            ////GetTifFile();

            // 終了
            this.Close();
        }

        private void GetOcrText()
        {
            try
            {
                if (System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif").Count() == 0)
                {
                    // 対象画像がないときは戻る
                    return;
                }

                DateTime dt = DateTime.Now;

                // ファイル名（日付時間部分）
                string fName = string.Format("{0:0000}", dt.Year) +
                               string.Format("{0:00}", dt.Month) +
                               string.Format("{0:00}", dt.Day) +
                               string.Format("{0:00}", dt.Hour) +
                               string.Format("{0:00}", dt.Minute) +
                               string.Format("{0:00}", dt.Second);

                int dNum = 0;

                //マルチTiff画像をシングルtifに分解する(SCANフォルダ → TRAYフォルダ)
                if (MultiTif_New(Properties.Settings.Default.scanPath, Properties.Settings.Default.workPath, Properties.Settings.Default.TrayPath))
                {
                    //WinReaderHandsのJOB名取得
                    string jobname = Properties.Settings.Default.WinReader_Job;

                    //ＯＣＲ対象画像ファイル数取得
                    int imgCnt = System.IO.Directory.GetFiles(Properties.Settings.Default.TrayPath).Count();

                    //WinReaderを起動して画像をスキャンしてOCR処理を実施する
                    WinReaderOCR(jobname);

                    // OCR認識結果ＣＳＶデータを出勤簿ごとに分割して画像ファイルと共にDATAフォルダへ移動する
                    LoadCsvDivide(fName, ref dNum, Properties.Settings.Default.dataPath);

                    //ログ表示
                    //logGridView(dataGridView1, imgCnt, "ＯＣＲ認識を行いました");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                //// ログ表示 : 2017/10/22
                //string msg = ex.Message + " 処理は継続されています。";
                //logGridView(dataGridView1, 0, msg);
            }
            finally
            {
                //notifyIcon1.Visible = true;
            }
        }


        ///------------------------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する：OpenCVバージョン</summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <param name="outPath">
        ///     分割後出力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool MultiTif_New(string scanPath, string InPath, string outPath)
        {
            string curFnm = "";

            try
            {
                Cursor = Cursors.WaitCursor;

                // 作業フォルダがなければ作成する
                if (!System.IO.Directory.Exists(InPath))
                {
                    System.IO.Directory.CreateDirectory(InPath);
                }

                // 出力先フォルダがなければ作成する
                if (!System.IO.Directory.Exists(outPath))
                {
                    System.IO.Directory.CreateDirectory(outPath);
                }

                // 作業フォルダ(WORK)内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
                foreach (string files in System.IO.Directory.GetFiles(InPath, "*"))
                {
                    System.IO.File.Delete(files);
                }

                // 出力先フォルダ(TRAY)内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
                foreach (string files in System.IO.Directory.GetFiles(outPath, "*"))
                {
                    System.IO.File.Delete(files);
                }

                // SCANフォルダ内の全てのファイルを作業フォルダに移動する
                foreach (string files in System.IO.Directory.GetFiles(scanPath, "*"))
                {
                    File.Copy(files, InPath + Path.GetFileName(files));
                    File.Delete(files);
                }

                int _pageCount = 0;
                string fnm = string.Empty;

                // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
                foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
                {
                    curFnm = files; 

                    //TIFFのImageCodecInfoを取得する//
                    ImageCodecInfo ici = GetEncoderInfo("image/tiff");

                    if (ici == null)
                    {
                        return false;
                    }

                    // 後片付けのためにusing 
                    using (System.IO.FileStream tifFS = new System.IO.FileStream(files, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        Image gim = Image.FromStream(tifFS);

                        FrameDimension gfd = new FrameDimension(gim.FrameDimensionsList[0]);

                        //全体のページ数を得る
                        int pageCount = gim.GetFrameCount(gfd);

                        for (int i = 0; i < pageCount; i++)
                        {
                            gim.SelectActiveFrame(gfd, i);

                            // 後片付けのためにusing
                            //using (Bitmap jj = new Bitmap(gim, 1728, 2322))　// コメント化 2021/11/05

                            // 後片付けのためにusing 2021/11/05
                            using (Bitmap jj = new Bitmap(gim, gim.Width, gim.Height))
                            {
                                // 画像解像度変更
                                jj.SetResolution(200F, 200F);

                                // ファイル名（日付時間部分）
                                string fName = string.Format("{0:0000}", DateTime.Today.Year) + string.Format("{0:00}", DateTime.Today.Month) +
                                               string.Format("{0:00}", DateTime.Today.Day)    + string.Format("{0:00}", DateTime.Now.Hour)    +
                                               string.Format("{0:00}", DateTime.Now.Minute)   + string.Format("{0:00}", DateTime.Now.Second);

                                _pageCount++;

                                // ファイル名設定
                                fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                                EncoderParameters ep = null;

                                // 圧縮方法を指定する（後片付けのためにusing）
                                using (ep = new EncoderParameters(1))
                                {
                                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

                                    // 画像保存
                                    jj.Save(fnm, ici, ep);
                                }
                            }
                        }

                        // 後片付け
                        gim.Dispose();
                    }
                }

                // InPathフォルダの全てのtifファイルを削除する
                foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
                {
                    System.IO.File.Delete(files);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(curFnm + " " + ex.Message);
                return false;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }


        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     FormOCRを起動してOCR処理を行う </summary>
        /// <param name="wrJobName">
        ///     FormOCRのJOB名</param>
        ///----------------------------------------------------------------------------------
        private void WinReaderOCR(string wrJobName)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                // WinReaderJOB起動文字列
                string JobName = @"""" + wrJobName + @"""" + " /H2";
                string winReader_exe = Properties.Settings.Default.WinReader_Path + Properties.Settings.Default.WinReader_Prg;

                // ProcessStartInfo の新しいインスタンスを生成する
                System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();

                // 起動するアプリケーションを設定する
                p.FileName = winReader_exe;

                // コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
                p.Arguments = JobName;

                // WinReaderを起動します
                System.Diagnostics.Process hProcess = System.Diagnostics.Process.Start(p);

                // taskが終了するまで待機する
                hProcess.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }


        ///-----------------------------------------------------------------
        /// <summary>
        ///     伝票ＣＳＶデータを一枚ごとに分割する </summary>
        /// <param name="dNum">
        ///     ファイル名末尾連番</param>
        /// <param name="cDir">
        ///     グループ別出力フォルダ</param>
        ///-----------------------------------------------------------------
        private void LoadCsvDivide(string fnm, ref int dNum, string cDir)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string imgName  = string.Empty;      // 画像ファイル名
                string firstFlg = global.flgOn.ToString();
                string[] stArrayData;               // CSVファイルを１行単位で格納する配列
                string newFnm   = string.Empty;
                int dCnt = 0;   // 処理件数

                // 対象ファイルの存在を確認します
                if (!System.IO.File.Exists(Properties.Settings.Default.readPath + Properties.Settings.Default.WinReader_OutFile))
                {
                    return;
                }

                //// 出力先フォルダ
                //string rPath = Properties.Settings.Default.dataPath + cDir + @"\";

                // 出力先フォルダが存在しないときは作成する
                if (!System.IO.Directory.Exists(cDir))
                {
                    System.IO.Directory.CreateDirectory(cDir);
                }

                // StreamReader の新しいインスタンスを生成する
                //入力ファイル
                System.IO.StreamReader inFile = new System.IO.StreamReader(Properties.Settings.Default.readPath + Properties.Settings.Default.WinReader_OutFile, Encoding.Default);

                // 読み込んだ結果をすべて格納するための変数を宣言する
                string stResult = string.Empty;
                string stBuffer;

                // 行番号
                int sRow = 0;

                // 読み込みできる文字がなくなるまで繰り返す
                while (inFile.Peek() >= 0)
                {
                    // ファイルを 1 行ずつ読み込む
                    stBuffer = inFile.ReadLine();

                    // カンマ区切りで分割して配列に格納する
                    stArrayData = stBuffer.Split(',');

                    //先頭に「*」があったら新たな伝票なのでCSVファイル作成
                    if ((stArrayData[0] == "*"))
                    {
                        //最初の伝票以外のとき
                        if (firstFlg != global.flgOn.ToString())
                        {
                            //ファイル書き出し
                            OutFileWrite(stResult, Properties.Settings.Default.readPath + imgName, cDir + newFnm);
                        }

                        firstFlg = global.flgOff.ToString();

                        // 伝票連番
                        dNum++;

                        // 処理件数
                        dCnt++;

                        // ファイル名
                        newFnm = fnm + dNum.ToString().PadLeft(3, '0');

                        //画像ファイル名を取得
                        imgName = stArrayData[1];

                        //文字列バッファをクリア
                        stResult = string.Empty;

                        // 文字列再校正（画像ファイル名を変更する）
                        stBuffer = string.Empty;
                        for (int i = 0; i < stArrayData.Length; i++)
                        {
                            if (stBuffer != string.Empty)
                            {
                                stBuffer += ",";
                            }

                            // 画像ファイル名を変更する
                            if (i == 1)
                            {
                                stArrayData[i] = newFnm + ".tif"; // 画像ファイル名を変更
                            }

                            // フィールド結合
                            stBuffer += stArrayData[i];
                        }

                        sRow = 0;
                    }
                    else
                    {
                        sRow++;
                    }

                    // 読み込んだものを追加で格納する
                    stResult += (stBuffer + Environment.NewLine);
                }

                // 後処理
                if (dNum > 0)
                {
                    //ファイル書き出し
                    OutFileWrite(stResult, Properties.Settings.Default.readPath + imgName, cDir + newFnm);

                    // 入力ファイルを閉じる
                    inFile.Close();

                    //入力ファイル削除 : "txtout.csv"
                    Utility.FileDelete(Properties.Settings.Default.readPath, Properties.Settings.Default.WinReader_OutFile);

                    //画像ファイル削除 : "WRH***.tif"
                    Utility.FileDelete(Properties.Settings.Default.readPath, "WRH*.tif");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     分割ファイルを書き出す </summary>
        /// <param name="tempResult">
        ///     書き出す文字列</param>
        /// <param name="tempImgName">
        ///     元画像ファイルパス</param>
        /// <param name="outFileName">
        ///     新ファイル名</param>
        ///----------------------------------------------------------------------------
        private void OutFileWrite(string tempResult, string tempImgName, string outFileName)
        {
            //出力ファイル 2017/11/20
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(outFileName + ".csv", false, System.Text.Encoding.GetEncoding(932));

            // ファイル書き出し
            outFile.Write(tempResult);

            //ファイルクローズ
            outFile.Close();

            //画像ファイルをコピー 2017/11/20
            System.IO.File.Copy(tempImgName, outFileName + ".tif");
        }


        ///-------------------------------------------------------------------------
        /// <summary>
        ///     MimeTypeで指定されたImageCodecInfoを探して返す </summary>
        /// <param name="mineType">
        ///     </param>
        /// <returns>
        ///     </returns>
        ///-------------------------------------------------------------------------
        private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo(string mineType)
        {
            //GDI+ に組み込まれたイメージ エンコーダに関する情報をすべて取得
            System.Drawing.Imaging.ImageCodecInfo[] encs = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();

            //指定されたMimeTypeを探して見つかれば返す
            foreach (System.Drawing.Imaging.ImageCodecInfo enc in encs)
            {
                if (enc.MimeType == mineType)
                {
                    return enc;
                }
            }
            return null;
        }



        private void GetTifFile()
        {
            ////////スキャン出力画像を確認
            //////string[] intif = System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif");
            //////if (intif.Length == 0)
            //////{
            //////    MessageBox.Show("ＯＣＲ変換処理対象の勤務記録表の画像ファイルが指定フォルダ " + Properties.Settings.Default.scanPath + " に存在しません", "スキャナ画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //////    return;
            //////}

            //////// READフォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            //////foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.readPath, "*"))
            //////{
            //////    System.IO.File.Delete(files);
            //////}

            //////// ファイル名を定義する
            //////string _fileName = Properties.Settings.Default.readPath + string.Format("{0:0000}", DateTime.Today.Year) +
            //////                                                  string.Format("{0:00}", DateTime.Today.Month) +
            //////                                                  string.Format("{0:00}", DateTime.Today.Day) +
            //////                                                  string.Format("{0:00}", DateTime.Now.Hour) +
            //////                                                  string.Format("{0:00}", DateTime.Now.Minute) +
            //////                                                  string.Format("{0:00}", DateTime.Now.Second);


            //////// tifファイルを認識する
            //////int _sNumber = 0;
            //////foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif"))
            //////{
            //////    //スキャナ出力先から作業フォルダへ移動する（同時にファイル名を書き換える）
            //////    _sNumber++;
            //////    System.IO.File.Move(files, _fileName + string.Format("{0:000}", _sNumber) + ".tif");
            //////}

            //////// OCR認識処理を実施する
            //////ocrMain(Properties.Settings.Default.readPath,
            //////        Properties.Settings.Default.ngPath, 
            //////        Properties.Settings.Default.dataPath, 
            //////        Properties.Settings.Default.fmtFilePath,
            //////        _sNumber);
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する </summary>
        /// <param name="InPath">
        ///     画像ファイルパス</param>
        ///---------------------------------------------------------------
        private void MultiTif(string InPath)
        {
            //スキャン出力画像を確認
            string[] intif = System.IO.Directory.GetFiles(InPath, "*.tif");
            if (intif.Length == 0)
            {
                MessageBox.Show("ＯＣＲ変換処理対象の出勤簿の画像ファイルが指定フォルダ " + Properties.Settings.Default.scanPath + " に存在しません", "スキャナ画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // READフォルダがなければ作成する
            if (System.IO.Directory.Exists(Properties.Settings.Default.readPath) == false)
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.readPath);

            // READフォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.readPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();
            _pageCount = 0;
            string fnm = string.Empty;

            // １．マルチTIFを分解して画像ファイルをREADフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
                    // ファイル名（日付時間部分）
                    string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                            string.Format("{0:00}", DateTime.Today.Month) +
                            string.Format("{0:00}", DateTime.Today.Day) +
                            string.Format("{0:00}", DateTime.Now.Hour) +
                            string.Format("{0:00}", DateTime.Now.Minute) +
                            string.Format("{0:00}", DateTime.Now.Second);

                    // ファイル名設定
                    _pageCount++;
                    fnm = Properties.Settings.Default.readPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.Tif, 0, i, i, 1, CodecsSavePageMode.Insert);
                }
            }

            // 2．InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     OCR処理を実施します </summary>
        /// <param name="InPath">
        ///     入力パス</param>
        /// <param name="NgPath">
        ///     NG出力パス</param>
        /// <param name="rePath">
        ///     OCR変換結果出力パス</param>
        /// <param name="FormatName">
        ///     書式ファイル名</param>
        /// <param name="fCnt">
        ///     書式ファイルの件数</param>
        ///---------------------------------------------------------------
        private void ocrMain(string InPath, string NgPath, string rePath, string FormatName, int fCnt)
        {
            IEngine en = null;		            // OCRエンジンのインスタンスを保持
            string ocr_csv = string.Empty;      // OCR変換出力CSVファイル
            int _okCount = 0;                   // OCR変換画像枚数
            int _ngCount = 0;                   // フォーマットアンマッチ画像枚数
            string fnm = string.Empty;          // ファイル名

            try
            {
                // 指定された出力先フォルダがなければ作成する
                if (System.IO.Directory.Exists(rePath) == false)
                    System.IO.Directory.CreateDirectory(rePath);

                // 指定されたNGの場合の出力先フォルダがなければ作成する
                if (System.IO.Directory.Exists(NgPath) == false)
                    System.IO.Directory.CreateDirectory(NgPath);

                // OCRエンジンのインスタンスの生成・取得
                en = EngineFactory.GetEngine();
                if (en == null)
                {
                    // エンジンが他で取得されている場合は、Release() されるまで取得できない
                    System.Console.WriteLine("SDKは使用中です");
                    return;
                }

                //オーナーフォームを無効にする
                this.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                IFormatList FormatList;
                IFormat Format;
                IField Field;
                int nPage;
                int ocrPage = 0;
                int fileCount = 0;

                // フォーマットのロード・設定
                FormatList = en.FormatList;
                FormatList.Add(FormatName);

                // tifファイルの認識
                foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
                {
                    nPage = 1;
                    while (true)
                    {
                        try
                        {
                            // 対象画像を設定する
                            en.SetBitmap(files, nPage);

                            //プログレスバー表示
                            fileCount++;
                            frmP.Text = "OCR変換処理実行中　" + fileCount.ToString() + "/" + fCnt.ToString();
                            frmP.progressValue = fileCount * 100 / fCnt;
                            frmP.ProgressStep();
                        }
                        catch (IDRException ex)
                        {
                            // ページ読み込みエラー
                            if (ex.No == ErrorCode.IDR_ERROR_FORM_FILEREAD)
                            {
                                // ページの終了
                                break;
                            }
                            else
                            {
                                // 例外のキャッチ
                                MessageBox.Show("例外が発生しました：Error No ={0:X}", ex.No.ToString());
                            }
                        }

                        //////Console.WriteLine("-----" + strImageFile + "の" + nPage + "ページ-----");
                        // 現在ロードされている画像を自動的に傾き補正する
                        en.AutoSkew();

                        // 傾き角度の取得
                        double angle = en.GetSkewAngle();
                        //////System.Console.WriteLine("時計回りに" + angle + "度傾き補正を行いました");

                        try
                        {
                            // 現在ロードされている画像を自動回転してマッチする番号を取得する
                            Format = en.MatchFormatRotate();
                            int direct = en.GetRotateAngle();

                            //画像ロード
                            RasterCodecs.Startup();
                            RasterCodecs cs = new RasterCodecs();
                            //RasterImage img;

                            // 描画時に使用される速度、品質、およびスタイルを制御します。 
                            //RasterPaintProperties prop = new RasterPaintProperties();
                            //prop = RasterPaintProperties.Default;
                            //prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                            //leadImg.PaintProperties = prop;

                            RasterImage img = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, 1);

                            RotateCommand rc = new RotateCommand();
                            rc.Angle = (direct) * 90 * 100;
                            rc.FillColor = new RasterColor(255, 255, 255);
                            rc.Flags = RotateCommandFlags.Resize;
                            rc.Run(img);
                            //rc.Run(leadImg.Image);

                            //cs.Save(leadImg.Image, files, RasterImageFormat.Tif, 0, 1, 1, 1, CodecsSavePageMode.Overwrite);
                            cs.Save(img, files, RasterImageFormat.Tif, 0, 1, 1, 1, CodecsSavePageMode.Overwrite);

                            // マッチしたフォーマットに登録されているフィールド数を取得
                            int fieldNum = Format.NumOfFields;
                            int matchNum = Format.FormatNo + 1;
                            //////System.Console.WriteLine(matchNum + "番目のフォーマットがマッチ");
                            int i = 1;
                            int fIndex = 0;
                            int dNum = 0;
                            ocr_csv = string.Empty;

                            // ファイルの先頭フィールドにファイル番号をセットします
                            ocr_csv = System.IO.Path.GetFileNameWithoutExtension(files) + ",";

                            // ファイルに画像ファイル名フィールドを付加します
                            ocr_csv += System.IO.Path.GetFileName(files) + ",";

                            // 認識されたフィールドを順次読み出します
                            Field = Format.Begin();
                            while (Field != null)
                            {
                                // 行先頭に日を付加
                                if (fIndex % 12 == 5)
                                {
                                    dNum++;
                                    ocr_csv += dNum.ToString() + ",";
                                }

                                // 指定フィールドを認識し、テキストを取得
                                string strText = Field.ExtractFieldText();
                                ocr_csv += strText;

                                // 改行付加
                                if (fIndex == 4 || (fIndex > 15 && ((fIndex - 4) % 12 == 0)))
                                    ocr_csv+=Environment.NewLine;
                                else ocr_csv += ",";    //カンマ付加

                                // 次のフィールドの取得
                                Field = Format.Next();
                                i += 1;

                                // フィールドインデックスインクリメント
                                fIndex++;
                            }

                            //出力ファイル
                            StreamWriter outFile = new StreamWriter(InPath + System.IO.Path.GetFileNameWithoutExtension(files) + ".csv", false, System.Text.Encoding.GetEncoding(932));
                            outFile.WriteLine(ocr_csv);
                            outFile.Close();

                            //OCR変換枚数カウント
                            _okCount++;     
                        }
                        catch (IDRWarning ex)
                        {
                            // Engine.MatchFormatRotate() で
                            // フォーマットにマッチしなかった場合の処理
                            if (ex.No == ErrorCode.IDR_WARN_FORM_NO_MATCH)
                            {
                                // NGフォルダへ移動する
                                File.Move(files, NgPath + System.IO.Path.GetFileName(files));
                                _ngCount++;　//NG枚数カウント
                            }
                        }

                        ocrPage++;
                        nPage += 1;
                    }
                }

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 終了メッセージ
                string finMessage = string.Empty;
                StringBuilder sb = new StringBuilder();
                sb.Append("OCR変換処理が終了しました");
                sb.Append(Environment.NewLine);
                sb.Append(Environment.NewLine);
                sb.Append("OK件数 : ");
                sb.Append(_okCount.ToString());
                sb.Append(Environment.NewLine);
                sb.Append("NG件数 : ");
                sb.Append(_ngCount.ToString());
                sb.Append(Environment.NewLine);

                MessageBox.Show(sb.ToString(), "処理終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // OCR変換画像とCSVデータをOCR結果出力フォルダへ移動する            
                foreach (string files in System.IO.Directory.GetFiles(InPath, "*.*"))
                {
                    File.Move(files, rePath + System.IO.Path.GetFileName(files));
                }

                FormatList.Delete(0);
            }
            catch (System.Exception ex)
            {
                // 例外のキャッチ
                string errMessage = string.Empty;
                errMessage += "System例外が発生しました：" + Environment.NewLine;
                errMessage += "必要なDLL等が実行モジュールと同ディレクトリに存在するか確認してください。：" + Environment.NewLine;
                errMessage += ex.Message.ToString();
                MessageBox.Show(errMessage, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                en.Release();
            }
        }

        private void frmOcr_Load(object sender, EventArgs e)
        {
            // フォーム最大表示設定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            // フォーム最小表示設定
            Utility.WindowsMinSize(this, this.Width, this.Height);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOcr_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
        
    }
}
