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

namespace insertForm {
    public partial class ForcefullyReadForm : Form {
        public ForcefullyReadForm() {
            InitializeComponent();
        }

        public static bool changeFlg = false;
        public static string DefaultPth= Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
        public static bool WorktxtFlg=false;
        public static int File_index_Number;
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        private void button2_Click(object sender, EventArgs e) {

            if (Form1.Downloaded == true) {
                txt.txtDelete(Program.Prev.AcPth);
                txt.txtDelete(Program.Prev.Ac);
                txt.txtDelete(Program.Prev.Tab);
                txt.txtDelete(Program.Prev.flp);
                txt.txtDelete(Program.Prev.lang);
            }

            Form1.Downloaded = false;

            if (listBoxTable.Text != "") {

                Program.SelectTablename = listBoxTable.Text;
                Form1.LookTable(Program.SelectAccessPath,Program.MainF.comboBox1);
            }
            this.Close();
        }

        #region クロージング！！！

        private void ForcefullyReadForm_FormClosing(object sender, FormClosingEventArgs e) {//閉じる時のイベント
            Program.MainF.Show();
            

            if (changeFlg==true) {
                Program.stdLang.std.Items.Clear();
                Program.stdLang.Lang.Items.Clear();

                if (txt.filechack(@"LanguageList.txt") == true) {
                    txt.TextRead(ref Program.stdLang.Lang, @"LanguageList.txt");
                } else {
                    MessageBox.Show("科目fileが見つかりません。\n終了します");
                    Application.Exit();
                }


            }


            Program.MainF.comboBox1.SelectedIndex = txt.match(Program.MainF.comboBox1, listBoxTable.Text);
        }

        #endregion


        private void button3_Click(object sender, EventArgs e) {
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

        private void button4_Click(object sender, EventArgs e) {
            //自分自身のフォームを最小化
            this.WindowState = FormWindowState.Minimized;
        }


        #region ダイアログ
        private void button1_Click(object sender, EventArgs e) {//----------------【　フォルダ選択　】
                                                                
            FolderBrowserDialog fbDialog = new FolderBrowserDialog();

            // ダイアログの説明文を指定する
            fbDialog.Description = "読み取るフォルダの選択";

            // デフォルトのフォルダを指定する
            fbDialog.SelectedPath = DefaultPth ;//ダイアログを開いた最初の場所


            // 「新しいフォルダーの作成する」ボタンを表示する
            fbDialog.ShowNewFolderButton = true;

            //フォルダを選択するダイアログを表示する
            if (fbDialog.ShowDialog() == DialogResult.OK&&fbDialog.SelectedPath=="") {
                MessageBox.Show(fbDialog.SelectedPath);

                string strp = fbDialog.SelectedPath;

                //----------------------------------------------//
                IEnumerable<string> files = System.IO.Directory.EnumerateFiles(strp, "*", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                //ファイルを列挙する
                foreach (string f in files) {
                    listBoxTxtPath0.Items.Add(f);
                    //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                    string txtname = Path.GetFileNameWithoutExtension(f);

                    listBoxTxtName00.Items.Add(txtname);


                }
                //--------------------------------------------------//

            } else {
                Console.WriteLine("キャンセルされました");
            }

            // オブジェクトを破棄する
            fbDialog.Dispose();

        }
        #endregion


        #region ロード！！！！！！

        private void ForcefullyReadForm_Load(object sender, EventArgs e) {//（　＾ω＾）ロード
            listBoxTxtName00.AllowDrop = true;//ドラッグドロップを可能にする

            // アイコンファイルのパス
            string p = Form1.FoldaList + @"usa1.ico";
            // パスを指定してアイコンのインスタンスを生成
            Icon icon = new Icon(p, 16, 16);
            // フォームのIconプロパティに設定
            this.Icon = icon;

         


                if (txt.filechack(Form1.FoldaList + @"LanguageList.txt") == true) {
                    txt.TextRead(ref LanguageList, Form1.FoldaList + @"LanguageList.txt");
                } else {
                    MessageBox.Show("科目fileが見つかりません。\n終了します");
                    Application.Exit();
                }

     
            
                AccessNamelbl.Text = Program.SelectAccessname;//選択されているAccessファイル名表示ラベル

            string AcPath = CONNECT_STRING + Program.SelectAccessPath;//AccessのPath

                //---------------------【テーブル名　取得　最初】----------------------------------//


                OleDbConnection conect = new OleDbConnection(AcPath);
                conect.Open();

                string[] opt = { null, null, null, "Table" };   //ユーザー定義のテーブルのみ取得を指定
                DataTable metaTable = conect.GetSchema("Tables", opt);   //データベースよりテーブル情報を取得
                DataColumn col = metaTable.Columns["TABLE_NAME"];


                foreach (DataRow row in metaTable.Rows) {//-------------
                    string tblName = row[col].ToString();//テーブル名

                    string sql = "SELECT * FROM " + tblName;//--SQL Select文

                    listBoxTable.Items.Add(tblName);//--テーブル名


                }//------------------【foreachの末尾】------------------

                conect.Close();
                conect.Dispose();

            //--------------------【テーブル名　取得　末尾】---------------------------------//


            #region //---前回の状態を呼び出し//--------------------
            if (Form1.Downloaded == true) {
                AccessNamelbl.Text = Program.Prev.dct["Access"];

                listBoxTable.Text = Program.Prev.dct["Table"];



                LanguageList.Text = Program.Prev.dct["lang"];

                Program.SelectAccessPath = Program.Prev.dct["AccessPath"];

            } else {
                listBoxTable.SelectedIndex = txt.match(listBoxTable, Program.SelectTablename);
                //LanguageList.SelectedIndex = 0;

 
                
            }
            #endregion
        }


        #endregion


        private void ForcefullyReadButton_Click(object sender, EventArgs e) {//----【 テキストファイル一斉読み込み 】
            if (listBoxTable.Text == ""|| LanguageList.Text=="") {

                MessageBox.Show("テーブル名または、言語名の未選択");
                return;
            }

            int count = 0;

            for (int i = 0; i <listBoxTxtPath0.Items.Count; i++) {
                listBoxTxtPath0.SelectedIndex = i;
                listBoxTxtName00.SelectedIndex = listBoxTxtPath0.SelectedIndex;



                textBox1.Text = TextRead(listBoxTxtPath0.Text,LanguageList);//TextRead() と textRead()




                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                txtnameBox.Text = Path.GetFileNameWithoutExtension(listBoxTxtName00.Text);



                Form1.strPath = Program.SelectAccessPath;

                //---------------------【テーブル名　取得　最初】----------------------------------//
                string AcPath = CONNECT_STRING + Form1.strPath;



                string tablePath;

                tablePath = listBoxTable.Text;//テーブル名は空白を入れると認識される文字が途切れてしまってエラーが起きるので、「_」を代わりに使用すること


           

                OleDbConnection conect = new OleDbConnection(AcPath);
                conect.Open();

                string sql = "INSERT INTO " + tablePath + " (言語,名前,サンプル) "   //ポイント
                    + "VALUES (@a,@b,@c)";

                OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand型変数

                string sample = "\r\n\r\n" + textBox1.Text;
                command.Parameters.AddWithValue("@a", LanguageList.Text);
                command.Parameters.AddWithValue("@b", txtnameBox.Text);
                command.Parameters.AddWithValue("@c", sample);
                //command.Parameters.AddWithValue("@d", text3.Text);


                int c = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

                //MessageBox.Show("追加：" + count);


                conect.Close();
                conect.Dispose();//ガベージコレクション（メモリの開放）
                count++;
                //LookTable();
            }
            MessageBox.Show("合計件数：" + count);
            listBoxTxtPath0.Items.Clear();
            listBoxTxtName00.Items.Clear();
            textBox1.Text = "";
        }

        //（　＾ω＾）・・・
        public static string TextRead(string Paths,ListBox langlist) {
            if (File.Exists(Paths)) {

                StreamReader sr;

                if (langlist.Text == "C言語") {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("shift_jis"));
                } else if (langlist.Text == "Java") {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("shift_jis"));
                } 
                else if (langlist.Text == "C++")//C言語のゲームサンプルの読み取りの場合
                {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("shift_jis"));
                } else if (langlist.Text == "JavaScript") {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("UTF-8"));
                } else {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("UTF-8"));
                }
                string text = sr.ReadToEnd();

                sr.Close();

                Console.Write(text);
                return text;
            } else {
                Console.WriteLine("ファイルが存在しません");
                return "失敗";
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) {

        }



        private void button6_Click(object sender, EventArgs e) {
            string s = Microsoft.VisualBasic.Interaction.InputBox("追加する文字列を入力してください");//インプットボックス使用
            if (txt.LangeListchack(LanguageList, s)) {//一致する文字列を探す※既に登録されていないか判定

            } else {
                txt.TxtStdWrite(ref LanguageList, s, Form1.FoldaList + "LanguageList.txt");
                LanguageList.Items.Add(s);
            }
        }


        private void SelectDeleteBtn_Click(object sender, EventArgs e) {
            if (listBoxTxtName00.Text!="") {


                int num = listBoxTxtName00.SelectedIndex;

                listBoxTxtName00.Items.RemoveAt(num);
                listBoxTxtPath0.Items.RemoveAt(num);

                if (num>=listBoxTxtPath0.Items.Count) {

                    listBoxTxtName00.SelectedIndex = listBoxTxtPath0.Items.Count-1;
                    listBoxTxtPath0.SelectedIndex = listBoxTxtName00.SelectedIndex;
                }
                else
                {
                    listBoxTxtName00.SelectedIndex = num;
                    listBoxTxtPath0.SelectedIndex = listBoxTxtName00.SelectedIndex;
                }
            } else {
                MessageBox.Show("削除するテキストファイルを選択してください。");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBoxTxtPath0.Items.Clear();
            listBoxTxtName00.Items.Clear();

        }


        private void listBoxTxtName00_DragEnter(object sender, DragEventArgs e) {
            e.Effect = DragDropEffects.All;
        }

        private void listBoxTxtName00_DragDrop(object sender, DragEventArgs e) {
            e.Effect = DragDropEffects.Copy;

            //コントロール内にドロップされたとき実行される
            //ドロップされたすべてのファイル名を取得する（フルパス）
            string[] fileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            List<string> FileDrop = new List<string>();


            FileDrop.AddRange(fileName);

            Drag_Drop.listCount(FileDrop, LanguageList, listBoxTxtPath0, listBoxTxtName00, NGWordbox);

            ////ListBoxに追加する
            //listBoxTxtName00.Items.AddRange(fileName);
        }

        private void listBoxTxtName00_MouseDoubleClick(object sender, MouseEventArgs e) {
            listBoxTxtPath0.SelectedIndex = listBoxTxtName00.SelectedIndex;

            Form1.memoName = "作業用";
            Form1.memotxt = txt.TextWorkRead(listBoxTxtPath0.Text,LanguageList);
            WorktxtFlg = true;

            Program.F2cntrl = new Form2();
            Program.F2cntrl.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e) {
            txt.TxtStdWrite(listBoxTable, AccessNamelbl, LanguageList);
            Application.Exit();
        }
    }


    }
