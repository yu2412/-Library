using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using insertForm;
using System.Windows.Forms;

namespace insertForm {
    class txt {

        

        public ComboBox std;
        public ListBox Lang;
        public ComboBox DaialogPth;
        public string[] filePath = new string[3];

        public static void Reading(ref TextBox T, string mess) {//登録用/


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamReader reader = new StreamReader(mess);


            // テキストを書き込む
            T.Text = reader.ReadToEnd();

            // ファイルを閉じる
            reader.Close();

        }






        #region 作業用テキスト1

        public static void Txt_OUT_Edit(string txt)//------　　　テキストサンプルを展開------------------------
        {
            txtDelete(Form1.FoldaList + "編集用テキスト.txt");

            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(Form1.FoldaList + @"編集用テキスト.txt", false, enc);


            // テキストを書き込む
            writer.WriteLine(txt);

            // ファイルを閉じる
            writer.Close();

        }//-----------------------------------//


        public static void Txt_EditEnd(ref TextBox T)//-----------Form１のテキスト取り込み用-取り込み内容の修正--修正処置------------------//
        {


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamReader reader = new StreamReader(Form1.FoldaList + @"編集用テキスト.txt");


            // テキストを書き込む
            T.Text = reader.ReadToEnd();

            // ファイルを閉じる
            reader.Close();

        }//-----------------------------

        public static void Txt_EditEnd(ref RichTextBox R)//------------アップデート用--終了処置------------------//
        {


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamReader reader = new StreamReader(Form1.FoldaList + @"編集用テキスト.txt");


            // テキストを書き込む
            R.Text = reader.ReadToEnd();

            // ファイルを閉じる
            reader.Close();

        }//-----------------------------

        #endregion

        #region 作業用テキスト2（複数のサンプルを展開して作業)
        public static void Txt_Multiple_Edit(string txtPath, string name, string S1)//-複数のサンプルを展開---
        {


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(txtPath + @"\" + name + ".txt", false, enc);//falseで上書き

            // テキストを書き込む
            writer.WriteLine(S1);//上書き内容（Accessへのパス）

            // ファイルを閉じる
            writer.Close();

        }





        public static void Txt_OUTEditEnd(string name)//-----１枚のサンプルを破棄---------------------------//
        {
            txtDelete( name + ".txt");

        }//-----------------------------

        #endregion
        public static void TextReadDF(string Paths,ref string[] dfst) {//Accessに取り込む時に使用
            if (File.Exists(Paths)) {

                StreamReader sr = new StreamReader(Paths, Encoding.GetEncoding("UTF-8"));

                dfst[0] = sr.ReadLine();
                dfst[1] = sr.ReadLine();
                sr.Close();


            } else {
                MessageBox.Show("ファイルが存在しません");
                return;
            }
        }
        public static string TextWorkRead(string Paths, ListBox langlist) {
            if (File.Exists(Paths)) {

                StreamReader sr;

                if (langlist.Text == "C言語") {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("shift_jis"));
                } else if (langlist.Text == "Java") {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("shift_jis"));
                } else if (langlist.Text == "C++")//C言語のゲームサンプルの読み取りの場合
                  {
                    sr = new StreamReader(Paths, Encoding.GetEncoding("shift_jis"));
                }
                  //else if (listLang.Text == "C++")
                  //{
                  //    sr = new StreamReader(Paths, Encoding.GetEncoding("UTF-8"));
                  //}
                  else {
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
        public static string TextRead(string Paths) {//Accessに取り込む時に使用
            if (File.Exists(Paths)) {

                StreamReader sr = new StreamReader(Paths, Encoding.GetEncoding("UTF-8"));

                string text = sr.ReadToEnd();

                sr.Close();

                Console.Write(text);
                return text;
            } else {
                Console.WriteLine("ファイルが存在しません");
                return "失敗";
            }
        }




        public static void TextRead(ref ListBox list, string Path) {//言語リストの読み込み


            if (File.Exists(Path)) {
                StreamReader sr = new StreamReader(Path, Encoding.GetEncoding("UTF-8"));

                while (sr.Peek() != -1) {
                    list.Items.Add(sr.ReadLine());
                }

                sr.Close();
                list.SelectedIndex = 0;
            } else {
                Console.WriteLine("ファイルが存在しません");

            }
        }

        public static bool PrevAccesRead(ref Dictionary< Ac,ComboBox> List) {//前回のセーブデータの確認

            if (File.Exists(Form1.FoldaList +@"AccessSave.txt")) {
                StreamReader sr = new StreamReader(Form1.FoldaList +@"AccessSave.txt", Encoding.GetEncoding("UTF-8"));

                try
                {
                    List[Ac.ACCESS_PATH].Text = sr.ReadLine();
                    List[Ac.ACCESS_NAME].Text = sr.ReadLine();
                    List[Ac.TABLE_NAME].Text = sr.ReadLine();

                    sr.Close();
                    //MessageBox.Show("発見");
                    if (filechack(List[Ac.ACCESS_PATH].Text) == false)
                    {

                        MessageBox.Show("エラーが起きました");
                        return false;

                    } 
                    return true;
                }
                catch(Exception e) 
                {
                    return false;
                }
            } else {
                //MessageBox.Show("エラーが起きました");
                return false;
            }
        }

        public static bool filechack(string fpth) {
            if (File.Exists(fpth)) {
                //StreamReader sr = new StreamReader(fpth, Encoding.GetEncoding("UTF-8"));
                //sr.Close();
                return true;
            } else {
                //MessageBox.Show("ありません");
                return false;
            }

        }

        public static bool filechacks(StringsFile filestr) {



            if (File.Exists(Form1.FoldaList+ @"PrevAcPath.txt") && File.Exists(Form1.FoldaList+@"PrevwDownloadAcName.txt") && File.Exists(Form1.FoldaList + @"PrevwDownloadTable.txt") && File.Exists(Form1.FoldaList + @"PrevwDownloadfilep.txt") && File.Exists(Form1.FoldaList + @"PrevwDownloadstd.txt") && File.Exists(Form1.FoldaList + @"PrevwDownloadLang.txt")) {

                StreamReader sr = new StreamReader(filestr.AcPth, Encoding.GetEncoding("UTF-8"));
                filestr.dct.Add("AccessPath", sr.ReadLine());
                sr.Close();


                StreamReader sr1 = new StreamReader(filestr.Ac, Encoding.GetEncoding("UTF-8"));
                filestr.dct.Add("Access",sr1.ReadLine());
                sr1.Close();


                StreamReader sr2 = new StreamReader(filestr.Tab, Encoding.GetEncoding("UTF-8"));
                filestr.dct.Add("Table", sr2.ReadLine());
                sr2.Close();

                StreamReader sr3 = new StreamReader(filestr.flp, Encoding.GetEncoding("UTF-8"));
                filestr.dct.Add("file", sr3.ReadLine());
                sr3.Close();


                StreamReader sr5 = new StreamReader(filestr.lang, Encoding.GetEncoding("UTF-8"));
                filestr.dct.Add("lang", sr5.ReadLine());
                sr5.Close();

                return true;
            } else {

                return false;
            }

        }


        public static void Save(string[] Ary) {//今回のセーブデータ作成
            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(Form1.FoldaList + @"AccessSave.txt", false, enc);//falseで上書き


            // テキストを書き込む
            writer.WriteLine(Ary[0]);//上書き内容Accessパス

            // ファイルを閉じる
            writer.Close();


            // ファイルを開く
            StreamWriter writer2 = new StreamWriter(Form1.FoldaList+"AccessSave.txt", true, enc);//追加

            // テキストを書き込む
            writer2.WriteLine(Ary[1]);//追加内容1//テーブル名
            // テキストを書き込む
            writer2.WriteLine(Ary[2]);//追加内容2Accessファイル名

            // ファイルを閉じる
            writer2.Close();

            //  MessageBox.Show("defaultを変更しました");

        }

        public static void TxtStdWrite(string txtPath,ComboBox cmb) {//登録用


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(txtPath, true, enc);//falseで上書き

            // テキストを書き込む
            writer.WriteLine(cmb.Text);//上書き内容（Accessへのパス）

            // ファイルを閉じる
            writer.Close();

        }

        public static void TxtStdWrite(ref ListBox Listb,string s,string txtPath) {//登録用


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(txtPath, true, enc);//falseで上書き

            // テキストを書き込む
            writer.WriteLine("");//改行用

            // テキストを書き込む
            writer.WriteLine(s);//

            // ファイルを閉じる
            writer.Close();

            Listb.Text = s;

        }

        public static void TxtStdWrite(ref ComboBox cmb, string s,string Path) {//登録用/


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(Path, true, enc);


            // テキストを書き込む
            writer.WriteLine(s);//

            // ファイルを閉じる
            writer.Close();

            cmb.Text = s;

        }
        public static void TxtStdWrite(ComboBox cmbstd,int selectnum) {//stdの今回使用ユーザー保存


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(Form1.FoldaList+@"defaultList.txt", false, enc);//falseで上書き

            // テキストを書き込む
            writer.WriteLine(cmbstd.SelectedIndex.ToString());//上書き内容（Accessへのパス）

            // ファイルを閉じる
            writer.Close();

            // ファイルを開く
            StreamWriter writer1 = new StreamWriter(Form1.FoldaList+@"defaultList.txt", true, enc);//falseで上書き

            // テキストを書き込む
            writer1.WriteLine(selectnum.ToString());//上書き内容（Accessへのパス）

            // ファイルを閉じる
            writer1.Close();

        }

        public static void TxtStdWrite(ListBox tab,Label label,ListBox lang) {//テキスト一斉保存フォームの履歴


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");



            // ファイルを開く
            StreamWriter writer = new StreamWriter(Form1.FoldaList+@"PrevwDownloadAcName.txt", false, enc);//falseで上書き
            // テキストを書き込む
            writer.WriteLine(label.Text);//上書き内容
            // ファイルを閉じる
            writer.Close();



            // ファイルを開く
            StreamWriter writer1 = new StreamWriter(Form1.FoldaList + @"PrevAcPath.txt", false, enc);//falseで上書き
            // テキストを書き込む
            writer1.WriteLine(Program.SelectAccessPath);//上書き内容
            // ファイルを閉じる
            writer1.Close();




            // ファイルを開く
            StreamWriter writer2 = new StreamWriter(Form1.FoldaList + @"PrevwDownloadTable.txt", false, enc);//falseで上書き
              writer2.WriteLine(tab.Text);//
            writer2.Close();



            // ファイルを開く
            StreamWriter writer6 = new StreamWriter(Form1.FoldaList + @"PrevwDownloadLang.txt", false, enc);//falseで上書き
            writer6.WriteLine(lang.Text);//
            writer6.Close();

            Application.Exit();

        }

        public static void TxtStdWrite(string txtPath, string S1)
        {//登録用


            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("UTF-8");

            // ファイルを開く
            StreamWriter writer = new StreamWriter(txtPath, true, enc);//falseで上書き

            // テキストを書き込む
            writer.WriteLine(S1);//上書き内容（Accessへのパス）

            // ファイルを閉じる
            writer.Close();

        }

        public static bool youserListchack(ComboBox combo, string key) {

            for (int i = 0; i < combo.Items.Count; i++) {

                //MessageBox.Show(combo.Text + "\n" + key);

                if (key == combo.Items[i].ToString())
                {
                    //MessageBox.Show("すでに登録されています");
                    return true;

                }

            }
            return false;

        }
        public static bool youserListchack(ComboBox combo, string key,string Touroku) {

            for (int i = 0; i < combo.Items.Count; i++) {

                //MessageBox.Show(combo.Text + "\n" + key);

                if (key == combo.Items[i].ToString()) {
                    MessageBox.Show("すでに登録されています");
                    return true;

                }

            }
            return false;

        }
        public static bool LangeListchack(ListBox list, string key) {

            for (int i = 0; i < list.Items.Count; i++) {

                //MessageBox.Show(combo.Text + "\n" + key);

                if (key == list.Items[i].ToString()) {
                    //MessageBox.Show("すでに登録されています");
                    return true;

                }

            }
            return false;

        }



        public static void txtDelete(string name) {
            File.Delete( name);//Form1.CurrentDirectorylibraryPth +
            //MessageBox.Show("削除完了");
        }

        public static bool strMacth(string S1, List<string> S2) {


            foreach (string A in S2) {
                if (A == S1) {
                    return true;
                }
            }
            return false;
        }

        public static bool strListMacth(string S1, ListBox list) {
            int prevNum = list.SelectedIndex;

            for (int i=0;i<list.Items.Count;i++) {
                if ( list.Items[i].ToString()== S1) {
                    return true;
                }
            }
            return false;
        }
        public static bool strListMacth(string S1, ComboBox comb) {
            int prevNum = comb.SelectedIndex;

            for (int i = 0; i < comb.Items.Count; i++) {
        
                if (comb.Items[i].ToString() == S1) {
                    return true;
                }
            }
            return false;
        }

        public static int match(ListBox listb, string key) {
            int i;

            for (i = 0; i < listb.Items.Count; i++) {

                if (key == listb.Items[i].ToString()) {
                    break;
                }
            }

            if (i >= listb.Items.Count) {
                return 0;
            } else {
                return i;
            }

        }
        public static int match(ComboBox comboBox, string key) {
            int i;
            int prev = comboBox.SelectedIndex;

            for (i = 0; i < comboBox.Items.Count; i++) {

                if (key == comboBox.Items[i].ToString()) {
                    break;
                }
            }

            if (i >= comboBox.Items.Count) {
                return prev;
            } else {
                return i;
            }

        }
        public static int match0(ComboBox comboBox, string key)
        {
            int i;

            for (i = 0; i < comboBox.Items.Count; i++)
            {
                if (key == comboBox.Items[i].ToString())
                {
                    break;
                }
            }

            if (i >= comboBox.Items.Count)
            {
                return 0;
            }
            else
            {
                return i;
            }

        }
        public static int match0(ListBox list1, string key)
        {
            int i;
            int prev = list1.SelectedIndex;

            for (i = 0; i < list1.Items.Count; i++)
            {
                if (key == list1.Items[i].ToString())
                {
                    break;
                }
            }

            if (i >= list1.Items.Count)
            {
                return 0;
            }
            else
            {
                return i;
            }

        }


        public static bool NG_Word_Chack(string Key)//エラーが起きるワードの入力チェック
        {
            switch (Key)
            {
                case "bit":return false;
                    
                case "Bit":
                    return false;
                    
                case "BIt":
                    return false;
                   
                case "BIT":
                    return false;
                 

                case "biT":
                    return false;
                 

                case "bIT":
                    return false;
              
                case "bIt":
                    return false;
             

                default:return true;
       
            }

        }



        public static bool Is_match(ListBox list, string key)
        {
            int i;
            int prev = list.SelectedIndex;

            for (i = 0; i < list.Items.Count; i++)
            {
                if (key == list.Items[i].ToString())
                {
                    break;
                }
            }

            if (i >= list.Items.Count)
            {
                return false;
            }
            else
            {
                return true;
            }

        }





    }
}
