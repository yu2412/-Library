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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        string File_supple = "Library_Supple";

        private void Form4_Load(object sender, EventArgs e)
        {
            
            string filePath = Path.GetFileNameWithoutExtension(Program.NewTable_AccessPath);
            textBoxAcname.Text= filePath;


            //---------------------【テーブル名　取得　最初】----------------------------------//
            string AcPath = CONNECT_STRING + Program.NewTable_AccessPath;
            


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

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string hante= textBox1.Text;

            List<string> S2 = new List<string>();
            S2.AddRange(listBox1.Items.Cast<string>().ToList());

           if(hante.Contains(" ")|| string.IsNullOrWhiteSpace(hante)) //string.IsNullOrWhiteSpace(hante))// null, 空文字, 半角スペース, 全角スペース,※hante.Contains("")は何を入力しても当てはまる
            {
                MessageBox.Show("テーブル名が未入力、または半角スペースが含まれています\n入力し直してください","入力ミス");
                return;
            }

            if (strMacth(hante,S2)==true)
            {
                MessageBox.Show("既存のテーブル名と被っております", "入力ミス");
                return;
            }


            DialogResult dr = MessageBox.Show("本当によろしいですか？", "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel)
            {
                MessageBox.Show("Cancelを押しました。");
                return;
            }
            else                //Cancel以外の動作
            {
                string tablePath;
                tablePath = textBox1.Text;//テーブル名は空白を入れると認識される文字が途切れてしまってエラーが起きるので、「_」を代わりに使用すること

                string AcPath = CONNECT_STRING + Program.SelectAccessPath;
                OleDbConnection conect = new OleDbConnection(AcPath);

                try
                {

                    conect.Open();

                    string sql = "CREATE TABLE " + tablePath + " (ID AUTOINCREMENT PRIMARY KEY, 言語 TEXT,名前 TEXT,サンプル TEXT,ポイント TEXT)"; //オートナンバー型に主キーの設定（PRIMARY KEY）//textBox1.Textは新規作成のテーブル名

            

                        OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand　命令を送る
                        int count = command.ExecuteNonQuery();//

                        #region --------------- テーブルの作成と同時にフォルダを作成------------------
                        Folder f = new Folder();

                        string a = Path_Cut.FileName(Program.SelectAccessPath);//Accessファイルへのパスからファイル名（拡張子なし）を取り出す

                      


                        string sup_path = Exe_Path.exe_Path() + File_supple + "\\" + a;
                        MessageBox.Show("Access名フォルダの存在確認" + sup_path);

                        f.Create_Folder_Exis(sup_path, tablePath);//Access名フォルダがなければ先に作成する＋テーブル名フォルダを作成

                        #endregion

                        MessageBox.Show("追加テーブル件数：" + count);
                    
                  
                }  catch (Exception e1) { MessageBox.Show(e1.Message); }
                conect.Close();
                conect.Dispose();//ガベージコレクション（メモリの開放）

                this.Close();
            }


        }

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.MainF.Show();
            Program.MainF.TableList();
        }

        public bool strMacth(string S1,List<string> S2)
        {
            

            foreach(string A in S2)
            {
                if (A==S1)
                {
                    return true;
                }
            }
            return false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
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
    }
}
