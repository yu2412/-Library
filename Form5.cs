using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using MyCreate;

namespace insertForm
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        string File_supple = "Library_Supple";

        private void Form5_Load(object sender, EventArgs e)
        {
            //---------------------【テーブル名　取得　最初】----------------------------------//
            string AcPath = CONNECT_STRING + Program.DaleteAccesePath;



            OleDbConnection conect = new OleDbConnection(AcPath);
            conect.Open();

            string[] opt = { null, null, null, "Table" };   //ユーザー定義のテーブルのみ取得を指定
            DataTable metaTable = conect.GetSchema("Tables", opt);   //データベースよりテーブル情報を取得
            DataColumn col = metaTable.Columns["TABLE_NAME"];


            foreach (DataRow row in metaTable.Rows) {//-------------
                string tblName = row[col].ToString();//テーブル名

                string sql = "SELECT * FROM " + tblName;//--SQL Select文

                listBox1.Items.Add(tblName);//--テーブル名

            }//------------------【foreachの末尾】------------------

            conect.Close();
            conect.Dispose();

            //--------------------【テーブル名　取得　末尾】-------------------------------

            if (Program.SelectTablename != "")
            {
                listBox1.SelectedIndex = txt.match0(listBox1, Program.SelectTablename);
            }
            else
            {
                listBox1.SelectedIndex = 0;
            }
           
        }

        private void button1_Click(object sender, EventArgs e)//テーブル削除
        {
            DialogResult dr = MessageBox.Show("テーブルを削除します。\n本当によろしいですか？\n"+"テーブル名："+listBox1.Text, "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel)
            {
                MessageBox.Show("Cancelを押しました。");
                return;
            }
            else//Cancel以外の動作
            {

                string tablePath;

                if (listBox1.Text=="") {
                    MessageBox.Show("テーブル名を選択してください");
                    return;
                }

                tablePath = listBox1.Text;//テーブル名は空白を入れると認識される文字が途切れてしまってエラーが起きるので、「_」を代わりに使用すること
                string AcPath = CONNECT_STRING + Program.DaleteAccesePath;

                OleDbConnection conect = new OleDbConnection(AcPath);
                conect.Open();


                string sql = "Drop Table " + listBox1.Text;

                try
                {

                    OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand　命令を送る



                    int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

                    
                   
                    string DeleTable = listBox1.Text;
                    string a= Path_Cut.FileName(Program.DaleteAccesePath);

                    Folder f = new Folder();
                    f.Delete_Folder(Exe_Path.exe_Path()+File_supple+"\\"+a+"\\"+DeleTable);

                    MessageBox.Show("テーブル削除：" + count);
                }
                catch (Exception e1) 
                {
                    MessageBox.Show(e1.Message);
                }

                conect.Close();
                conect.Dispose();//ガベージコレクション（メモリの開放）
            }


            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //自分自身のフォームの状態を調べる
            switch (this.WindowState)
            {
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

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.MainF.Show();

            Program.MainF.TableList();
            Program.MainF.comboBox0.SelectedIndex = 0;

            Program.MainF.TableList();
            
            Form1.LookTable(Program.DaleteAccesePath, Program.MainF.comboBox1);
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("テーブルを削除します。\n本当によろしいですか？\n" + "テーブル名：" + listBox1.Text, "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel)
            {
                MessageBox.Show("Cancelを押しました。");
                return;
            }
            else//Cancel以外の動作
            {

                string tablePath;

                if (listBox1.Text == "")
                {
                    MessageBox.Show("テーブル名を選択してください");
                    return;
                }

                tablePath = listBox1.Text;//テーブル名は空白を入れると認識される文字が途切れてしまってエラーが起きるので、「_」を代わりに使用すること
                string AcPath = CONNECT_STRING + Program.DaleteAccesePath;

                OleDbConnection conect = new OleDbConnection(AcPath);
                conect.Open();


                string sql = "Drop Table " + listBox1.Text;


                OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand　命令を送る



                int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

                MessageBox.Show("テーブル削除：" + count);


                conect.Close();
                conect.Dispose();//ガベージコレクション（メモリの開放）
            }


            this.Close();
        }
    }
}
