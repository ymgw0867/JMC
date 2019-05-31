using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using JMC.Common;

namespace JMC.Config
{
    public partial class frmCalendar : Form
    {
        public frmCalendar()
        {
            InitializeComponent();
        }

        private void frmCalendar_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            GridViewSetting(dataGridView1); //グリッドビュー設定
            ComboYear();                    // 対象年コンボボックス値セット
            GridViewShow(dataGridView1);    // グリッドビュー表示
            DispClr();                      // 画面初期化

            // 休日コンボボックス値セット
            Utility.comboHoliday.Load(comboBox1);
        }

        // ID
        string _ID;

        // 登録モード
        int _fMode = 0;

        // グリッドビューカラム名
        private string cDate = "c1";
        private string cGekkyu = "c2";
        private string cJikyu = "c3";
        private string cMemo = "c4";
        private string cID = "c5";

        /// <summary>
        /// グリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 322;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cDate, "年月日");
                tempDGV.Columns.Add(cMemo, "名称");
                tempDGV.Columns.Add(cGekkyu, "社員");
                tempDGV.Columns.Add(cJikyu, "パート");
                tempDGV.Columns.Add(cID, "ID");
                tempDGV.Columns[cID].Visible = false;

                tempDGV.Columns[cDate].Width = 110;
                tempDGV.Columns[cMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[cGekkyu].Width = 80;
                tempDGV.Columns[cJikyu].Width = 80;

                tempDGV.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cGekkyu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cJikyu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

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

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 休日対象年
        /// </summary>
        private void ComboYear()
        {
            comboBox2.Items.Clear();

            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            // コンボボックスに登録されている休日の年をセットする
            string mySql = "select distinct year(年月日) as y from 休日";
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = mySql;
            dr = sCom.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr["y"].ToString());
            }
            dr.Close();
            sCom.Connection.Close();

            // 今年を初期表示とする
            comboBox2.SelectedIndex = -1;
            for (int i = 0; i < comboBox2.Items.Count; i++)
			{
                if (comboBox2.Items[i].ToString() == DateTime.Today.Year.ToString())
                {
                    comboBox2.SelectedIndex = i;
                    break;
                }
			}

            // 当年の休日が登録されていないときは一番最近の年を初期表示とする
            if (comboBox2.Items.Count != 0)
            {
                if (comboBox2.SelectedIndex == -1) comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            }
        }
        
        /// <summary>
        /// 休日データをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewShow(DataGridView tempGrid)
        {
            if (comboBox2.Text != string.Empty)
            {
                //データベース接続
                SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                OleDbDataReader dr;

                string mySql = "select * from 休日";
                mySql += " where year(年月日) = " + comboBox2.Text;
                mySql += " order by 年月日";

                sCom.Connection = Con.cnOpen();
                sCom.CommandText = mySql;
                dr = sCom.ExecuteReader();

                int iX = 0;
                tempGrid.RowCount = 0;

                while (dr.Read())
                {
                    tempGrid.Rows.Add();

                    tempGrid[cDate, iX].Value = DateTime.Parse(dr["年月日"].ToString()).ToShortDateString();
                    tempGrid[cMemo, iX].Value = dr["名称"].ToString();

                    if (dr["月給者"].ToString() == "0")
                        tempGrid[cGekkyu, iX].Value = string.Empty;
                    else tempGrid[cGekkyu, iX].Value = "○";

                    if (dr["時給者"].ToString() == "0")
                        tempGrid[cJikyu, iX].Value = string.Empty;
                    else tempGrid[cJikyu, iX].Value = "○";

                    tempGrid[cID, iX].Value = dr["ID"].ToString();

                    iX++;
                }

                dr.Close();
                sCom.Connection.Close();

                tempGrid.CurrentCell = null;
            }
        }

        private void DispClr()
        {
            txtDate.Text = string.Empty;
            comboBox1.Text = string.Empty;
            checkBox1.Checked = false;
            checkBox2.Checked = false;

            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            btnClr.Enabled = false;
            monthCalendar1.Enabled = true;

            _fMode = 0;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            txtDate.Text = monthCalendar1.SelectionStart.ToString("yyyy/MM/dd (ddd)");

            // 休日名称を表示
            string md = monthCalendar1.SelectionStart.ToString("MM/dd");
            Utility.comboHoliday.selectedIndex(comboBox1, md);


            checkBox1.Checked = true;
            checkBox2.Checked = true;
            btnUpdate.Enabled = true;
            btnClr.Enabled = true;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (txtDate.Text == string.Empty)
            {
                MessageBox.Show("日付が選択されていません", "休日設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            switch (_fMode)
            {
                case 0:
                    if (!dataSearch(monthCalendar1.SelectionStart))
                    {
                        if (MessageBox.Show(txtDate.Text + " を登録しますか？", "休日登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                        dataInsert(monthCalendar1.SelectionStart);
                    }
                    else
                    {
                        MessageBox.Show("既に登録済みの日付です", "休日設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    break;
                case 1:
                    if (MessageBox.Show(txtDate.Text + " を更新しますか？", "休日登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                    dataUpdate(DateTime.Parse(txtDate.Text));
                    break;
                default:
                    break;
            }

            ComboYear();
            GridViewShow(dataGridView1);
            DispClr();
        }

        /// <summary>
        /// 休日テーブルに休日データを新規に登録する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        private void dataInsert(DateTime dt)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("insert into 休日 (年月日,名称,月給者,時給者,更新年月日) ");
            sb.Append("values (?,?,?,?,?) ");
            
            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@Date", dt.ToShortDateString());
            sCom.Parameters.AddWithValue("@memo", comboBox1.Text);

            int gs = 0;
            if (checkBox1.Checked) gs = 1;
            sCom.Parameters.AddWithValue("@Gekkyu", gs);

            int js = 0;
            if (checkBox2.Checked) js = 1;
            sCom.Parameters.AddWithValue("@Jikyu", js);

            sCom.Parameters.AddWithValue("@date2", DateTime.Today);
            
            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }

        /// <summary>
        /// 休日データを更新する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        private void dataUpdate(DateTime dt)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("update 休日 set ");
            sb.Append("年月日=?,名称=?,月給者=?,時給者=?,更新年月日=? ");
            sb.Append("where ID=?");

            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@date", txtDate.Text.Substring(0, 10));
            sCom.Parameters.AddWithValue("@memo", comboBox1.Text);

            int gs = 0;
            if (checkBox1.Checked) gs = 1;
            sCom.Parameters.AddWithValue("@Gekkyu", gs);

            int Js = 0;
            if (checkBox2.Checked) Js = 1;
            sCom.Parameters.AddWithValue("@Jikyu", Js);

            sCom.Parameters.AddWithValue("@date2", DateTime.Today);
            sCom.Parameters.AddWithValue("@id", _ID);

            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }

        /// <summary>
        /// 休日データを削除する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        private void dataDelete(DateTime dt)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("delete from 休日 ");
            sb.Append("where ID=?");

            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@id", _ID);

            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        /// <summary>
        /// グリッドビューの選択された行データを表示する
        /// </summary>
        /// <param name="g">データグリッドビューオブジェクト</param>
        private void GetGridViewData(DataGridView g)
        {
            if (g.SelectedRows.Count == 0) return;

            int r = g.SelectedRows[0].Index;

            string y = g[cDate, r].Value.ToString();

            txtDate.Text = DateTime.Parse(y).ToString("yyyy/MM/dd (ddd)");
            comboBox1.Text = g[cMemo, r].Value.ToString();

            if (g[cGekkyu, r].Value.ToString() == "○")
                checkBox1.Checked = true;
            else checkBox1.Checked = false;

            if (g[cJikyu, r].Value.ToString() == "○")
                checkBox2.Checked = true;
            else checkBox2.Checked = false;

            _ID = g[cID, r].Value.ToString();

            btnUpdate.Enabled = true;
            btnDelete.Enabled = true;
            btnClr.Enabled = true;
            //monthCalendar1.Enabled = false;
            _fMode = 1;
        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            DispClr();
        }

        /// <summary>
        /// データ削除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(txtDate.Text + " を削除してよろしいですか？", "休日削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
            dataDelete(DateTime.Parse(txtDate.Text));

            ComboYear();
            GridViewShow(dataGridView1);
            DispClr();
        }

        /// <summary>
        /// 休日データを検索する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        /// <returns>true:データあり、false:データなし</returns>
        private bool dataSearch(DateTime dt)
        {
            bool _Rtn = false;

            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select * from 休日 ");
            sb.Append("where 年月日=?");

            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@Date", dt);

            OleDbDataReader dR;
            dR = sCom.ExecuteReader();

            if (dR.HasRows) _Rtn = true;
            dR.Close();
            sCom.Connection.Close();

            return _Rtn;
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCalendar_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            GetGridViewData(dataGridView1);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridViewShow(dataGridView1);
        }
    }
}
