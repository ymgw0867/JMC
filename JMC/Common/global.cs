using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JMC
{
    class global
    {
        public static string pblImageFile;

        //表示関係
        public static float miMdlZoomRate = 0;      //現在の表示倍率

        //表示倍率（%）
        public static float ZOOM_RATE = 0.23f;      // 標準倍率：300dpi
        public static float ZOOM_RATE200 = 0.34f;   // 標準倍率：200dpi
        public static float ZOOM_MAX = 2.00f;       // 最大倍率
        public static float ZOOM_MIN = 0.05f;       // 最小倍率
        public static float ZOOM_STEP = 0.02f;      // ステップ倍率
        public static float ZOOM_NOW;               // 現在の倍率

        public static int RECTD_NOW;                // 現在の座標
        public static int RECTS_NOW;                // 現在の座標
        public static int RECT_STEP = 20;           // ステップ座標

        //和暦西暦変換
        public const int rekiCnv = 1988;    //西暦、和暦変換

        //エラーチェック関連
        public static string errID;         //エラーデータID
        public static int errNumber;        //エラー項目番号
        public static int errRow;           //エラー行
        public static string errMsg;        //エラーメッセージ

        //エラー項目番号
        public const int eNothing = 0;      //エラーなし
        public const int eYearMonth = 1;    //対象年月
        public const int eMonth = 2;       //対象月
        public const int eShainNo = 3;      //個人番号
        public const int eShozoku = 4;     //所属コード
        public const int eDay = 6;          //日
        public const int eTokubetsu = 7;    //休暇記号
        public const int eYukyu = 8;        //有給休暇
        public const int eSH = 9;            // 開始時
        public const int eSM = 10;          // 開始分
        public const int eEH = 11;          // 終了時
        public const int eEM = 12;          // 終了分
        public const int eKKH = 13;          // 規定内休憩・時間
        public const int eKKM = 14;          // 規定内休憩・分
        public const int eKSH = 15;          // 深夜帯休憩・時間
        public const int eKSM = 16;          // 深夜帯休憩・分
        public const int eTH = 17;          // 実働時
        public const int eTM = 18;          // 実働分
        public const int eNoCheck = 19;     // 未チェック出勤簿
        public const int e20 = 20;          //20時
        public const int e21 = 21;          //21時
        public const int e22 = 22;          //22時
        public const int eTotal = 23;       //日別勤務合計
        public const int eMTotal = 24;      //月間勤務合計
        public const int eYkigou1 = 25;     //役職記号1
        public const int eYkigou2 = 26;     //役職記号2
        public const int eYkigou3 = 27;     //役職記号3
        public const int eYkigou4 = 28;     //役職記号4
        public const int eYkigou5 = 29;     //役職記号5
        public const int eYkTeate = 30;     //役職手当合計
        public const int eSoTeate = 31;     //その他手当合計
        public const int eKtTeate = 32;     //交通費合計
        public const int eGTeate = 33;      //総合計
        public const int eKINMU_KUBUN = 34; // 勤務先区分

        //汎用データファイル名
        public static string OKFILE = "勤怠データ"; 

        //汎用データヘッダ項目
        public const string H1 = @"""EBAS001""";
        public const string H2 = @"""LTLT001""";
        public const string H3 = @"""LTLT003""";
        public const string H4 = @"""LTLT004""";
        public const string H5 = @"""LTDT001""";
        public const string H6 = @"""LTDT002""";
        public const string H7 = @"""LTDT003""";
        public const string H8 = @"""LTDT004""";

        //ローカルMDB関連
        public const string MDBFILE = "en_ocr.mdb";         //MDBファイル名
        public const string MDBTEMP = "en_ocr_Temp.mdb";    //最適化一時ファイル名
        public const string MDBBACK = "en_ocr_Back.mdb";    //最適化後バックアップファイル名

        public static int flgOn = 1;        //フラグ有り(1)
        public static int flgOff = 0;       //フラグなし(0)


        public static int TOUGETSU = 0;     //当月扱い
        public static int YOKUGETSU = 1;    //翌月扱い

        //ＯＣＲ処理ＣＳＶデータの検証要素
        public static int CSVLENGTH = 197;          //データフィールド数 2011/06/11
        public static int CSVFILENAMELENGTH = 21;   //ファイル名の文字数 2011/06/11  
 
        // 勤務記録表
        public static int STARTTIME = 8;    // 単位記入開始時間帯
        public static int ENDTIME = 22;     // 単位記入終了時間帯
        public static int TANNIMAX = 4;     // 単位最大値
        public static int WEEKLIMIT = 160;  // 週労働時間基準単位：40時間
        public static int DAYLIMIT = 32;    // 一日あたり労働時間基準単位：8時間

        // 対象年月
        //public static int pblYear;      // 対象年
        //public static int pblMonth;     // 対象月

        public static int ShozokuLength = 5;        // 所属コード桁数
        public static int ShainLength = 5;          // 社員コード桁数
        public static int ShozokuMaxLength = 5;     // 所属コードＭＡＸ桁数
        public static int ShainMaxLength = 5;       // 社員コードＭＡＸ桁数

        // 休暇記号
        public static string TOKUBETSU_KYUKA = "1";     // 特別休暇
        public static string KEKKIN_KYUKA = "2";        // 欠勤
        public static string CHISOU_KYUKA = "3";        // 遅刻・早退
        public static string FURIKYU_KYUKA = "4";       // 振替休日
        public static string FURIDE_KYUKA = "5";        // 振替出勤

        // 有給記号
        public static string ZENNICHI_YUKYU = "0";      // 全日有給
        public static string HANNICHI_YUKYU = "1";      // 半日有給
        public static string H1_YUKYU = "1";            // 1H有給
        public static string H2_YUKYU = "2";            // 2H有給
        public static string H3_YUKYU = "3";            // 3H有給
        public static string H4_YUKYU = "4";            // 4H有給
        public static string H5_YUKYU = "5";            // 5H有給
        public static string H6_YUKYU = "6";            // 6H有給
        public static string H7_YUKYU = "7";            // 7H有給

        // 深夜時間帯
        public static DateTime dt2200 = DateTime.Parse("22:00");
        public static DateTime dt0500 = DateTime.Parse("05:00");
        public static DateTime dt0800 = DateTime.Parse("08:00");

        // ChangeValueStatus
        public static bool ChangeValueStatus = true;

        public const int STATUS_SHAIN = 1;
        public const int STATUS_PART = 2;
        public const int FORM_ADDMODE = 0;
        public const int FORM_EDITMODE = 1;

        // 勤務先区分
        public const int KINMU_MAIN = 1;
        public const int KINMU_SUB = 2;

        // プレ印刷モード
        public const int CSV_MODE = 1;
        public const int XLS_MODE = 2;
    }
}
