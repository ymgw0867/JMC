using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using FujiXerox.DocuWorks.Toolkit;

namespace JMC.OCR
{
    public class xdwCreate
    {
        /// <summary>
        /// イメージファイルからdocuworkファイルを出力する
        /// </summary>
        /// <param name="inPath">
        /// 入力イメージファイルパス
        /// </param>
        /// <param name="outPath">
        /// 出力docuworkファイルパス
        /// </param>
        /// 
        public static void FromTiff(string inPath, string outPath)
        {
            // docuwork出力オプション
            Xdwapi.XDW_CREATE_OPTION_EX2 op = new Xdwapi.XDW_CREATE_OPTION_EX2();
            op.FitImage = Xdwapi.XDW_CREATE_FIT;

            // docuworkファイル出力
            int api_result = Xdwapi.XDW_CreateXdwFromImageFile(inPath, outPath, op);
            if (api_result < 0) MessageBox.Show("docuworkファイルの出力に失敗しました。");
        }
    }
}
