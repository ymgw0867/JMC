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
        /// DataControl�N���X�̊�{�N���X
        /// </summary>
        public class BaseControl
        {
            private DBConnect DBConnect;
            protected OleDbConnection dbControlCn;

            // BaseControl�̃R���X�g���N�^�BDBConnect�N���X�̃C���X�^���X���쐬���܂��B
            public BaseControl(string dbName)
            {
                // �f�[�^�x�[�X���I�[�v������
                DBConnect = new DBConnect(dbName);
            }

            // �f�[�^�x�[�X�ɐڑ����R�l�N�V��������Ԃ�
            public OleDbConnection GetConnection()
            {
                dbControlCn = DBConnect.Cn;
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            // �f�[�^�R���g���[���N���X�̃R���X�g���N�^
            public DataControl(string dbName):base(dbName)
            {
            }

            /// <summary>
            /// �f�[�^�x�[�X�ڑ�����
            /// </summary>
            public void Close()
            {
                if (dbControlCn.State == ConnectionState.Open)
                {
                    dbControlCn.Close();
                }
            }

            /// <summary>
            /// �C�ӂ�SQL�����s����
            /// </summary>
            /// <param name="tempSql">SQL��</param>
            /// <returns>���� : true, ���s : false</returns>
            public bool FreeSql(string tempSql)
            {
                bool rValue = false;

                try
                {
                    OleDbCommand sCom = new OleDbCommand();
                    sCom.CommandText = tempSql;
                    sCom.Connection = GetConnection();

                    //SQL�̎��s
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
            /// �f�[�^���[�_�[���擾����
            /// </summary>
            /// <param name="tempSQL">SQL��</param>
            /// <returns>�f�[�^���[�_�[</returns>
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
        /// SQLServer�f�[�^�x�[�X�ڑ��N���X
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
                    // MySeting���ڂ̎擾
                    sServerName = Properties.Settings.Default.SQLServerName;    // �T�[�o��
                    sLogin = Properties.Settings.Default.SQLLogin;              // ���O�C����
                    sPass = Properties.Settings.Default.SQLPass;                // �p�X���[�h
                    sDatabase = dbName;                                         // �f�[�^�x�[�X��

                    // �f�[�^�x�[�X�ڑ�������
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
