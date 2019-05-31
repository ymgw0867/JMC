using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using JMC.Common;

namespace JMC.Common
{
    class dbControl
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private DBConnect DBConnect;
            protected OleDbConnection dbControlCn;

            // BaseControlのコンストラクタ。DBConnectクラスのインスタンスを作成します。
            public BaseControl(string dbName)
            {
                // データベースをオープンする
                DBConnect = new DBConnect(dbName);
            }

            // データベースに接続しコネクション情報を返す
            public OleDbConnection GetConnection()
            {
                dbControlCn = DBConnect.Cn;
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            // データコントロールクラスのコンストラクタ
            public DataControl(string dbName):base(dbName)
            {
            }

            /// <summary>
            /// データベース接続解除
            /// </summary>
            public void Close()
            {
                if (dbControlCn.State == ConnectionState.Open)
                {
                    dbControlCn.Close();
                }
            }

            /// <summary>
            /// 任意のSQLを実行する
            /// </summary>
            /// <param name="tempSql">SQL文</param>
            /// <returns>成功 : true, 失敗 : false</returns>
            public bool FreeSql(string tempSql)
            {
                bool rValue = false;

                try
                {
                    OleDbCommand sCom = new OleDbCommand();
                    sCom.CommandText = tempSql;
                    sCom.Connection = GetConnection();

                    //SQLの実行
                    sCom.ExecuteNonQuery();
                    rValue = true;
                }
                catch
                {
                    rValue = false;
                }

                return rValue;
            }

            /// <summary>
            /// データリーダーを取得する
            /// </summary>
            /// <param name="tempSQL">SQL文</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FreeReader(string tempSQL)
            {
                OleDbCommand sCom = new OleDbCommand();
                sCom.CommandText = tempSQL;
                sCom.Connection = GetConnection();
                OleDbDataReader dR = sCom.ExecuteReader();

                return dR;
            }
        }

        /// <summary>
        /// SQLServerデータベース接続クラス
        /// </summary>
        public class DBConnect
        {
            OleDbConnection cn = new OleDbConnection();

            public OleDbConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            private string sServerName;
            private string sLogin;
            private string sPass;
            private string sDatabase;

            public DBConnect(string dbName)
            {
                try
                {
                    // MySeting項目の取得
                    sServerName = Properties.Settings.Default.SQLServerName;    // サーバ名
                    sLogin = Properties.Settings.Default.SQLLogin;              // ログイン名
                    sPass = Properties.Settings.Default.SQLPass;                // パスワード
                    sDatabase = dbName;                                         // データベース名

                    // データベース接続文字列
                    cn.ConnectionString = "";
                    cn.ConnectionString += "Provider=SQLOLEDB;";
                    cn.ConnectionString += "SERVER=" + sServerName + ";";
                    cn.ConnectionString += "DataBase=" + sDatabase + ";";
                    cn.ConnectionString += "UID=" + sLogin + ";";
                    cn.ConnectionString += "PWD=" + sPass + ";";
                    cn.ConnectionString += "WSID=";

                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }
    }
}
