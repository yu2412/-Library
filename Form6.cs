using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using MyCreate;

namespace insertForm {

    delegate void DelegCalc();//デリゲート

    public partial class Form6 : Form {
        public Form6() {
            InitializeComponent();
        }
        //接続文字列
        public const string CONNECT_STRING = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source= ";
        public static List<string> AcPathList;

        DelegCalc TableListView;

        public Form6(List<string> AcPath)
        {
            InitializeComponent();
            Form6.AcPathList = AcPath;
        }


 

        #region Load　ロード！！！！！！！！！

        private void Form6_Load(object sender, EventArgs e) {

        

            // アイコンファイルのパス
            string p = Form1.FoldaList + @"F6.ico";
            // パスを指定してアイコンのインスタンスを生成
            Icon icon = new Icon(p, 16, 16);
            // フォームのIconプロパティに設定
            this.Icon = icon;

        }
        //--------------------【ロード　末尾】---------------
        #endregion



        #region -------------【 Shownイベント   】---------------------------------------//

        private void Form6_Shown(object sender, EventArgs e)
        {
            comboBox0.Items.AddRange(Form6.AcPathList.ToArray());
            for (int i = 0; i < comboBox0.Items.Count; i++)
            {
                listBox00.Items.Add(Path.GetFileNameWithoutExtension(comboBox0.Items[i].ToString()));//拡張子無し
            }

            comboBox0.SelectedIndex = comboBox0.Items.Count > 0 ? 0 : comboBox0.SelectedIndex;

            listBox00.SelectedIndex = comboBox0.SelectedIndex;

            TableList();
        }

        //------------    【   末尾  】     ----------------//
        #endregion





        #region -------------【 テーブル名一覧   】---------------------------------------//
        private void TableList() 
        {
            //---------------------【テーブル名　取得　最初】----------------------------------//
              listBox1.Items.Clear();
            List<string> TableName = new List<string>();
            Access_SQL_Copy.TableList2(comboBox0.Text, ref TableName);

            listBox1.Items.AddRange(TableName.ToArray());
            if ( listBox1.Items.Count> 0)
            {
         
                listBox1.SelectedIndex = 0;
            }
    
        }
        //------------    【   末尾  】     ----------------//
        #endregion


        private void button1_Click(object sender, EventArgs e) {

                Program.SelectTablename = listBox1.Text;
                Program.SelectAccessPath = comboBox0.Text;
                Form1.LookTableKEY = true;
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e) {
            //自分自身のフォームの状態を調べる
            switch (this.WindowState) {
                case FormWindowState.Normal:
                    //MessageBox.Show("普通の状態です");
                    //自分自身のフォームを最大化
                    this.WindowState = FormWindowState.Maximized;
                    break;

                case FormWindowState.Maximized:
                    //MessageBox.Show("最大化されています");
                    //自分自身のフォームを最大化
                    this.WindowState = FormWindowState.Normal;
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e) {
            //自分自身のフォームを最小
            this.WindowState = FormWindowState.Minimized;
        }

        private void button2_Click(object sender, EventArgs e) {
            Form1.LookTableKEY = false;
            this.Close();
        }

        #region -------------【  CLOSEイベント  】---------------------------------------//
        private void Form6_FormClosing(object sender, FormClosingEventArgs e) {//////////------------------閉じているときの作業

                Program.MainF.Show();

                Program.MainF.comboBox0.Items.Clear();
                Program.MainF.comboBox00.Items.Clear();
                Program.MainF.comboBox1.Items.Clear();
                Program.MainF.comboBox0.Items.AddRange(comboBox0.Items.Cast<string>().ToArray());
                Program.MainF.comboBox00.Items.AddRange(listBox00.Items.Cast<string>().ToArray());
                Program.MainF.comboBox00.SelectedIndex = txt.match0(Program.MainF.comboBox00,listBox00.Text);
                Program.MainF.comboBox1.SelectedIndex  = txt.match0(Program.MainF.comboBox1, listBox1.Text);

        }
        //------------    【   末尾  】     ----------------//
        #endregion

        private void listBox00_SelectedIndexChanged(object sender, EventArgs e) {
            comboBox0.SelectedIndex = listBox00.SelectedIndex;//選択順番を連動させる
            listBox1.Items.Clear();

            ListBox lis = (ListBox)sender;

            if (lis.Text != "")
            {
                TableListView = TableList;
            }
            else
            {
                TableListView = ErrorMesg;
            }
            TableListView();

        }



        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e) {
            Program.SelectTablename = listBox1.Text;

            Program.SelectTablename = listBox1.Text;
            Program.SelectAccessname = listBox00.Text;
            Program.SelectAccessPath = comboBox0.Text;
            Form1.LookTableKEY = true;

            this.Close();
        }

        public bool strMacth(string S1, List<string> S2) {

            foreach (string A in S2) {
                if (A == S1) {
                    return true;
                }
            }
            return false;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox00_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void ErrorMesg() 
        {
            MessageBox.Show("想定外のエラー発生：selectIndexChange");
        }

    }
}
