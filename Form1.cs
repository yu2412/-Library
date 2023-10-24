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
using insertForm;
using Microsoft.VisualBasic;
using MyCreate;
using Microsoft.VisualBasic.FileIO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace insertForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        static string File_supple = "Library_Supple";
        public static List<string> AcPathList;
        public static List<string> AcNameList;
        public static Size size1 = new Size(978, 786);
       public static Size size2 = new Size(400, 90);

        #region グローバル

        public static string KeyWord;
        public static DataGridView SearchResultView;
        List<string> ColumName = new List<string> { "言語", "名前", "サンプル" };

        public static int C = 0;
        public static string GlobalPath;
        public static bool MultiWorking = false;

        public static string[,] updFrm3;

        public static System.Windows.Forms.TextBox[] Box;

        public static string strPath;

        public static string memotxt;
        public static string memoName;

        public static TabControl tbCntl;
        public static DataGridView DaGrVu1;
        public static Dictionary<Ac, System.Windows.Forms.ComboBox> CombList;
        public static System.Windows.Forms.Button[] ActButtons;
        public static System.Windows.Forms.ComboBox YouserBox;
        public static bool PrevFlg = false;

        public static string[] ComboBoxAcces;
        public static string tablename;
        public static bool Downloaded = false;
        public static int[] DefaultNumber;

        public static System.Windows.Forms.Label Labe2;

        public static string ExcellPath = @"\List";

        public static bool LookTableKEY = false;//エラーの防止用

        public static string NewTable;
        public static string CurrentDirectorylibraryPth;
        public static string Folda;
	    public static string FoldaList;
        public static int File_index_Number;

        static string[] strLine = new string[500];
        static string[] TestLine = new string[] { "aaa", "bbb", "ccc", "ddd" };
        #endregion



        #region //----------【　Load　ロード！！】--------------------//
        private void Form1_Load(object sender, EventArgs e)
        {
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.ShowIcon = true;
            this.ShowInTaskbar = true;
            string Jpath = System.Windows.Forms.Application.ExecutablePath;
            string folderPath = System.IO.Path.GetDirectoryName(Jpath);
            Folda = folderPath + @"\";
	        FoldaList=Folda+@"List\";
            Program.MulWork = new MultiList();
            Program.stdLang = new txt();
            Program.Prev = new StringsFile();
            Program.stdLang.Lang = LangList;

            if (txt.filechack(FoldaList+@"LanguageList.txt") == true)
            {
                txt.TextRead(ref LangList, FoldaList+@"LanguageList.txt");
            }
            else
            {
                MessageBox.Show("科目fileが見つかりません。\n終了します");
            }

            //--------------------------------------（　＾ω＾）・・・
            string path = System.Windows.Forms.Application.ExecutablePath;
            string folderPath1 = Path.GetDirectoryName(path);
            CurrentDirectorylibraryPth = folderPath1 + @"\library\";//登録情報のテキストファイルへのパス

            //--------------------------------------（　＾ω＾）・・・

            Program.NewTable_AccessPath = "";//新規作成のテーブル用

            Program.SelectTableList = new List<string>();//テーブル選択用
            Form1.AcPathList = new List<string>();//Accessへのパス
            Form1.AcNameList = new List<string>();//Accessファイル名のみ

            Program.daGrVi1 = dataGridView1;

            CombList = new Dictionary<Ac, System.Windows.Forms.ComboBox> { { Ac.ACCESS_PATH, comboBox0 }, { Ac.ACCESS_NAME, comboBox00 }, { Ac.TABLE_NAME, comboBox1 } };
                                                                            //Accessパス//Accessの名//テーブル名

            //----------------------------------------------//
            IEnumerable<string> files = System.IO.Directory.EnumerateFiles(folderPath1, "*.accdb", System.IO.SearchOption.AllDirectories);

            //ファイルを列挙する
            foreach (string f in files)
            {
                AcPathList.Add(f);
                //comboBox0.Items.Add(f);

            }
            //--------------------------------------------------//

            //ファイルを列挙する
            foreach (string f in files)
            {
                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                string filePath = Path.GetFileNameWithoutExtension(f);
                AcNameList.Add(filePath);
                //comboBox00.Items.Add(filePath);
            }
            //--------------------------------------------------//

            comboBox0.Items.AddRange(AcPathList.ToArray());
            comboBox00.Items.AddRange(AcNameList.ToArray());

            //comboBox0.Text = folderPath + @"\アルゴリズム集.accdb";
            //comboBox0.SelectedIndex = 0;//1番下の選択肢を選択
            //-----------------------------------------------------------------------------------//
            TableList();
            //-----------------------------------------------------------------------------//
            // フォーム右上の最大化／最小化／閉じる各ボタンを非表示に（全部使用不可）
           // this.ControlBox = false;
            // フォームのアイコンを表示する
            this.ShowIcon = true;
            updFrm3 = new string[2, 5];
            DaGrVu1 = dataGridView1;
            tbCntl = tabControl1;
            DefaultNumber = new int[2];
            Box = new System.Windows.Forms.TextBox[9];
            Box[0] = Text1;
            Box[1] = textBox1;
            Labe2 = label2;
            ActButtons = new System.Windows.Forms.Button[2];
        
            ActButtons[0] = button5;
            ActButtons[1] = button6;

            //Program.combos = new ComboBox[3];

            dataGridView1.ColumnCount = 5;//dataSourceから値を入れないのであれば必要
            //表示
           // dataGridView1.Rows.Add();//行の生成

            dataGridView1.Columns[0].HeaderText = "ID";
            dataGridView1.Columns[1].HeaderText = "言語";
            dataGridView1.Columns[2].HeaderText = "名前";
            dataGridView1.Columns[3].HeaderText = "サンプル";
            dataGridView1.Columns[4].HeaderText = "ポイント";

            //DataGridView1でセル、行、列が複数選択されないようにする【行選択させるのに大事！！】
            dataGridView1.MultiSelect = true;

            Program.AccessRecord = new string[100, 1];
            Program.Save = new string[3];//前回の最後の行動の記録
            Program.Save[0] = null;

            //------------------------（　＾ω＾）・・・//前回一斉取り込みデータがあるかどうかを確認
            if (txt.filechacks(Program.Prev))
            {
                Downloaded = true;
            }

            if (txt.filechack(FoldaList+"defaultList.txt"))
            {//前回のForm1の作業データがあるかどうか
                string[] df = new string[2];
                txt.TextReadDF(FoldaList+"defaultList.txt", ref df);
            }
            else
            {
                File_index_Number=filePointbox00.SelectedIndex = 0;
            }

            if (txt.filechack(FoldaList + "AccessSave.txt"))
            {//前回のForm1の作業データがあるかどうか

                if (txt.PrevAccesRead(ref CombList))
                {
                    PrevFlg = true;


                    Program.SelectTablename = comboBox1.Text;
                    Program.SelectAccessPath = comboBox0.Text;
                    Program.SelectAccessname = comboBox00.Text;
                }

            }


        }//----------【　ロードイベントの末尾】ーーーーーーーーーーーーーーーーーーー
        #endregion

        #region -------------【    Showイベント 】---------------------------------

        private void Form1_Shown(object sender, EventArgs e)
        {//（　＾ω＾）・・・【Shownイベント！！！】

            if (Downloaded == true)
            {
                Program.SelectAccessPath = comboBox0.Text;
                Program.SelectAccessname = comboBox00.Text;
                ForcefullyReadOpen();
            }
            else if (PrevFlg == false)
            {//前回使用データがあるかどうか                                             //Program.SelectTableList.AddRange(comboBox1.Items.Cast<string>().ToList());


                SelectTableForm6();

            }
            else
            {
                LookTable(Program.SelectAccessPath, comboBox1);
                // アプリケーションの設定を読み込む
                Properties.Settings.Default.Reload();

                try {
                    LangList.SelectedIndex = Properties.Settings.Default.UseLang;//【設定した名前が"Name"】
            
                }
                catch (Exception) {

                }
                //----------------------------------------//
            }
        }
        #endregion

        #region -------------挿入Query_1件追加---------------
        private void InsertConnectbtn_Click(object sender, EventArgs e)
        {
            if (LookTableKEY == false)
            {
                MessageBox.Show("テーブルを選択してください。");
                return;
            }
            List<string> Qu_Val = new List<string> { {LangList.Text },{textBox1.Text },{Text1.Text } };


            if (Access_SQL_Copy.INSERT_Query(comboBox0.Text, comboBox1.Text, ColumName, Qu_Val)) 
            {
                MessageBox.Show("追加成功");
                LookTable(comboBox0.Text, comboBox1);
                tabControl1.SelectedIndex = 0;
            }
            else 
            {
                MessageBox.Show("入力失敗");
            }

        }
        #endregion

        #region Form6へ移動ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            SelectTableForm6();
        }



        public void SelectTableForm6()
        {
            if (comboBox0.Text != "")
            {
                Program.SelectAccessPath = comboBox0.Text;
                AcPathList.Clear();
                AcPathList.AddRange(comboBox0.Items.Cast<string>().ToArray());
            }

            Program.F6cntrl = new Form6(AcPathList);
            Program.F6cntrl.Show();
            Program.MainF.Hide();
        }
        #endregion



        private void 閉じるToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Application.Exit();
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                return;
            }
            else
            {
                tabControl1.SelectedIndex -= 1;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                tabControl1.SelectedIndex = 0;
            }
            else
            {
                tabControl1.SelectedIndex = 1;
            }

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Form1 f = new Form1();

            this.dataGridView1.Font = new System.Drawing.Font("Arial", (float)numericUpDown1.Value);
        }

        #region テーブルの指定した項目行を削除するクエリ
        private void button5_Click(object sender, EventArgs e)
        {//////------------------------------削除ボタン
            if (LookTableKEY == false)
            {
                MessageBox.Show("テーブルを選択してください。");
                return;
            }

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("データがありません");
                return;
            }

            this.Size = size2;//縮小化

            //現在のセルの行インデックスを格納


            List<int> nRow = new List<int>();

            //選択されているすべての項目インデックス番号を取得する
            //dataGridView1.CurrentCell.RowIndex;
            string moji = "";
            foreach (DataGridViewRow d in dataGridView1.SelectedRows)//選択された行のみ
            {
                nRow.Add(d.Index);
                moji += "\r\n" + "ID番号：" + dataGridView1.Rows[d.Index].Cells[0].Value.ToString();
            }
                DialogResult dr = MessageBox.Show("本当によろしいですか？\n" +moji , "確認", MessageBoxButtons.OKCancel);       
        if (dr == System.Windows.Forms.DialogResult.OK) {

            bool Flg = true;
            for (int i=0;i<nRow.Count;i++) {

                
                    int Rowint = nRow[i];
                    string key = dataGridView1.Rows[Rowint].Cells[0].Value.ToString();

                    if (Access_SQL_Copy.Delete_Query(comboBox0.Text, comboBox1.Text, key)) 
                    {  

                    } 
                    else { Flg = false; }
            }
            if (Flg) {LookTable(comboBox0.Text, comboBox1); }
        }

            MessageBox.Show("処理完了(≧ω≦)");
            this.Size = size1;//
        }
        #endregion

        private void 最大化ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Size = (this.Size.Height <= size2.Height) ? size1 : size2;
        }

        private void 最小化ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;
        }

        private void Access選択toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(comboBox0.Text+ "に別のAccessファイルから追加を開始します", "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel) {
                //  MessageBox.Show("Cancelを押しました。");
                return;
            }

            Program.MulWork.nextWorks.Add(new FormNextWork());
            Program.MulWork.nextWorks[0].PrevAccessPath = comboBox0.Text;
            Program.MulWork.nextWorks[0].PrevAccessName = comboBox00.Text;
            Program.MulWork.nextWorks[0].PrevTableName = comboBox1.Text;

            Program.SelectTablename = comboBox1.Text;
                Program.SelectTableList.Clear();
                foreach (var item in comboBox1.Items) 
                {
                    Program.SelectTableList.Add(item.ToString());
                }

                Form1.AcPathList.Clear();

                foreach (var item in comboBox0.Items) 
                {
                    AcPathList.Add(item.ToString());
                }

                NewMoveForm moveForm = new NewMoveForm(Program.SelectTableList);
                Program.MainF.Hide();
                DialogResult result= moveForm.ShowDialog();

                if (result == DialogResult.OK) 
                {
                    LookTable(Program.SelectAccessPath,comboBox1);
                }
        }

        #region     ----------------【更新フォーム】
        private void button6_Click(object sender, EventArgs e)//更新用ファームへ移動
        {


            if (LookTableKEY == false)
            {
                MessageBox.Show("テーブルを選択してください。");
                return;
            }

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("データがありません");
                return;
            }

            int RowNow = dataGridView1.CurrentRow.Index;//選択番号
            this.Size = size2;//縮小化
            DialogResult dr = MessageBox.Show("ID番号：" + dataGridView1.Rows[RowNow].Cells[0].Value.ToString() + "\n更新画面に移ります", "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルしました。");
                this.Size = size1;//
                return;
            }
            else//"Cancel以外の動作"で実行
            {
                tablename = comboBox1.Text;

                Program.F3cntrl = new Form3();

                // カラム名を指定
                updFrm3[0, 0] = dataGridView1.Columns[0].HeaderText;//IDの番号
                updFrm3[0, 1] = dataGridView1.Columns[1].HeaderText; ;//言語
                updFrm3[0, 2] = dataGridView1.Columns[2].HeaderText;//名前
                updFrm3[0, 3] = dataGridView1.Columns[3].HeaderText;//サンプル
                updFrm3[0, 4] = dataGridView1.Columns[4].HeaderText;//ポイント

                updFrm3[1, 0] = dataGridView1.Rows[RowNow].Cells[0].Value.ToString();//IDの番号
                updFrm3[1, 1] = dataGridView1.Rows[RowNow].Cells[1].Value.ToString();//言語
                updFrm3[1, 2] = dataGridView1.Rows[RowNow].Cells[2].Value.ToString();//名前
                updFrm3[1, 3] = dataGridView1.Rows[RowNow].Cells[3].Value.ToString();//サンプル
                updFrm3[1, 4] = dataGridView1.Rows[RowNow].Cells[4].Value.ToString();//ポイント

                UPdateForm();
            }
        }

        public void UPdateForm()
        {
            Program.SelectAccessPath = comboBox0.Text;

            Program.F3cntrl.Show();
            Program.MainF.Hide();
        }

        #endregion

        public void NewTableForm()
        {
            Program.F4cntrl.Show();
            Program.MainF.Hide();
        }


        #region テーブルの削除フォーム
        private void テーブル削除toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //string DeTable = comboBox00.Text;
            string DeTable = Program.SelectAccessPath;

            DialogResult dr = MessageBox.Show($"【{DeTable}】のテーブルの削除画面に移動します", "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルしました。");
                return;
            }
            else//"Cancel以外の動作"で実行
            {
                //string[] s = new string[comboBox1.Items.Count];

                Program.DaleteAccesePath = Program.SelectAccessPath;

                if (comboBox1.Text != "")
                {
                    Program.SelectTablename = comboBox1.Text;
                }
                else
                {
                    Program.SelectTablename = "";
                }

                Program.F5cntrl = new Form5();
                DeleteTableForm();
                //テーブル一覧のデータをリストに格納↓コンボ１とすり替える



            }
        }


        public void DeleteTableForm()
        {

            Program.F5cntrl.Show();
            Program.MainF.Hide();
        }
        #endregion

        #region テーブルの新規作成用の、Form４へ移動
        private void テーブルの新規作成ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewTable = Program.SelectAccessPath;

            DialogResult dr = MessageBox.Show($"【{NewTable}】の新規テーブルを作成します", "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルしました。");
                return;
            }
            else//"Cancel以外の動作"で実行
            {
                //string[] s = new string[comboBox1.Items.Count];
                //テーブル一覧のデータをリストに格納↓コンボ１とすり替える
                Program.NewTable_AccessPath = comboBox0.Text;

                Program.F4cntrl = new Form4();
                NewTableForm();

            }
        }
        #endregion






        #region Accessのファイル名コンボボックスの選択が変更した時のイベント
        private void comboBox00_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox0.SelectedIndex = comboBox00.SelectedIndex;

            TableList();
        }

        #endregion

        #region テーブル名のコンボボックスの選択が変更されたイベント
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            List<List<string>> result = new List<List<string>>();

            Access_SQL_Copy.Select_Query2(comboBox0.Text,comboBox1.Text,ref result);
            dataGridView1.ColumnCount = 5;

           // MessageBox.Show("行数："+result.Count);

            dataGridView1.Rows.Clear();
            for (int i = 0; i < result.Count; i++)
            {
                dataGridView1.Rows.Add();

                for (int j = 0; j < result[i].Count; j++)
                {
                        dataGridView1.Rows[i].Cells[j].Value = result[i][j];
                }
            }

            dataGridView1.RowTemplate.Height = 50;
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 100;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 500;
            dataGridView1.Columns[4].Width = 500;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ReadOnly = true;
            tabControl1.SelectedIndex = 0;

            LookTableKEY = true;
            button5.Visible = true;
            button6.Visible = true;
        }
        #endregion


        #region クロージングメソッド（フォームを閉じてる最中）
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {//フォームを閉じている途中で今の開いてる状態を保存するか確認
            bool SaveData = false;

            DialogResult dr = MessageBox.Show("今回の作業テーブルを保存しますか？", "確認", MessageBoxButtons.OKCancel);
            if (dr != System.Windows.Forms.DialogResult.Cancel)
            {
                SaveData = true;
            }

            if (SaveData)
            {
                Program.Save = new string[3];
                Program.Save[0] = comboBox0.Text;
                Program.Save[1] = comboBox00.Text;
                Program.Save[2] = comboBox1.Text;

                txt.Save(Program.Save);



            }
            else
            {


                //txt.txtDelete(@"defaultList.txt");
                txt.txtDelete(FoldaList+ @"AccessSave.txt");
            }

            try {
                Properties.Settings.Default.UseLang = LangList.SelectedIndex;//【保存したい値を代入】
            }
            catch (Exception) 
            {

            }

            // アプリケーションの設定を保存する
            Properties.Settings.Default.Save();
        }
        #endregion


        public void ForcefullyReadOpen()
        {
            Program.ForcefullyReadCntrl = new ForcefullyReadForm();

            Program.ForcefullyReadCntrl.Show();
            Program.MainF.Hide();
        }//テキストファイル一斉読み込みフォーム


        private void button7_Click(object sender, EventArgs e)
        {

            string s = Microsoft.VisualBasic.Interaction.InputBox("追加する文字列を入力してください");//インプットボックス使用
            if (txt.LangeListchack(LangList, s))
            {//一致する文字列を探す※既に登録されていないか判定

            }
            else
            {
                txt.TxtStdWrite(ref LangList, s, FoldaList +"LanguageList.txt");
                LangList.Items.Add(s);
            }
        }

        private void filePointbox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //-------------------------------------------------------------------------------------------//

        #region -------------【 テーブル名一覧   】---------------------------------------//
       public void TableList()
       {
            bool Connect_Flg = false;
            //---------------------【テーブル名　取得　最初】----------------------------------//
            comboBox1.Items.Clear();

            List<string> TableName = new List<string>();
            Access_SQL_Copy.TableList2(comboBox0.Text, ref TableName);

            comboBox1.Items.AddRange(TableName.ToArray());

            if (comboBox1.Items.Count > 0)
            {
                Connect_Flg = true;
                comboBox1.SelectedIndex = 0;
            }
            else
            {
                dataGridView1.Rows.Clear();
                Connect_Flg = false; 
            }
        }
        //------------    【   末尾  】     ----------------//
        #endregion



        #region テーブルの中身_展開_コンボボックス！！
        public static void LookTable(string Ac, System.Windows.Forms.ComboBox TableNa)
        {//----------------------------------------------------------------------【　テーブルの項目全て表示一覧　】

            bool Connect_Flg = false;

            List<List<string>> result = new List<List<string>>();

            Access_SQL_Copy.Select_Query2(Ac, TableNa.Text, ref result);
            DaGrVu1.Rows.Clear();
            for (int i = 0; i < result.Count; i++)
            {
                DaGrVu1.Rows.Add();

                for (int j = 0; j < result[i].Count; j++)
                {
                    DaGrVu1.Rows[i].Cells[j].Value = result[i][j];
                }
            }

            if (Connect_Flg)
            {

                LookTableKEY = true;
                ActButtons[0].Visible = true;
                ActButtons[1].Visible = true;
                ActButtons[2].Visible = true;
            }
        }//LookTable  テーブルの中身全て一覧！！！！

        #endregion

        //-----------------------------------------------------//

        #region テーブルの中身_展開_リストボックス！！

        public static void LookTable(string Ac, System.Windows.Forms.ListBox TableNa)
        {//----------------------------------------------------------------------【　テーブルの項目全て表示一覧　】
            bool Connect_Flg = false;

            List<List<string>> result = new List<List<string>>();

            Access_SQL_Copy.Select_Query2(Ac,TableNa.Text,ref result);

            for (int i = 0; i < result.Count; i++)
            {
                DaGrVu1.Rows.Add();

                for (int j = 0; j < result[i].Count; j++)
                {
                    DaGrVu1.Rows[i].Cells[j].Value = result[i][j];
                }
            }

            try
            {
                DaGrVu1.RowTemplate.Height = 50;
                DaGrVu1.Columns[0].Width = 40;
                DaGrVu1.Columns[1].Width = 100;
                DaGrVu1.Columns[2].Width = 200;
                DaGrVu1.Columns[3].Width = 700;
                DaGrVu1.Columns[4].Width = 500;
                //dataGridView1.Columns[5].Width = 80;
                //dataGridView1.Columns[6].Width = 50;
                tbCntl.SelectedIndex = 0;//tabControl1.SelectedIndex = 0;
                Connect_Flg = true;

            }catch(Exception e) { MessageBox.Show(e.Message); }

            if (Connect_Flg)
            {

                LookTableKEY = true;
                ActButtons[0].Visible = true;
                ActButtons[1].Visible = true;
                ActButtons[2].Visible = true;

            }
        }//LookTable  テーブルの中身全て一覧！！！！


        #endregion


        #region テーブルの名前変更！！

        public void TableNameEdit(System.Windows.Forms.ComboBox comb1, System.Windows.Forms.ComboBox comb0)
        {
           
            string PrevTableName = comb1.Text;

            string s = Microsoft.VisualBasic.Interaction.InputBox(comb1.Text + " 　テーブルの名前を変更します");//インプットボックス使用
            if (txt.strListMacth(s, comb1)||txt.NG_Word_Chack(s)==false)  //一致する文字列を探す※既に登録されていないか判定
            {
                MessageBox.Show("すでに同じ名前があります");
                return;
            }
            else
            {
                DialogResult dr = MessageBox.Show("テーブル名を変更します。\n本当によろしいですか？\n" + "テーブル名：" + PrevTableName + "\r\n変更後：" + s, "確認", MessageBoxButtons.OKCancel);
                if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    MessageBox.Show("Cancelを押しました。");
                    return;
                }
                else//Cancel以外の動作
                {

                    Program.SelectAccessPath = comb0.Text;
                    Program.SelectTablename = comb1.Text;

                    TableNameChengeSQL(Program.SelectAccessPath, Program.SelectTablename, s);

                   Access_SQL_Copy.Delete_Table(comboBox0.Text,comboBox1.Text );

                    TableList();
                    comb1.SelectedIndex = txt.match(comb1, s);
                    Program.SelectTablename = comb1.Text;
                    LookTable(Program.SelectAccessPath, comb1);
                }
            }
        }
        #endregion


        #region テーブル複製メソッド＿コピーを別名で作成

        public static void TableNameChengeSQL(string Ac, string TaNa, string NextName)
        {

            //---------------------【テーブル名　取得　最初】----------------------------------//


            try { 
                MessageBox.Show("テーブル複製：");
             
                string a = Path_Cut.FileName(Ac);

                string path1 = Exe_Path.exe_Path() + File_supple + "\\" + a + "\\"+ TaNa;
                string path2 = Exe_Path.exe_Path() + File_supple + "\\" + a + "\\" + NextName;

                Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(path1, path2);
            }
            catch(Exception e) 
            {
                MessageBox.Show(e.Message);
            }


        }

        #endregion


        private void button8_Click(object sender, EventArgs e)
        {
            //メモ帳起動
            System.Diagnostics.Process p = System.Diagnostics.Process.Start("notepad.exe");

            //アイドル状態まで待機
            p.WaitForInputIdle();


        }

        private void button9_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.IDataObject data = Clipboard.GetDataObject();
            if (data != null)
            {

                if (data.GetDataPresent(typeof(string)))
                {
                    Text1.Text = Clipboard.GetText();//クリップボードゲット

                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("String型に変換出来ません");
                }
            }
        }

        private void テーブル名変更toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (comboBox0.Text == "")
            {
                MessageBox.Show("Accessのパスが指定されていません");
                return;
            }
            this.Size = size2;//サイズ
            TableNameEdit(comboBox1, comboBox0);
            this.Size = size1;//サイズをもどす


        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                // MessageBox.Show("データがありません");
                return;
            }

            if ((string)dataGridView1.CurrentCell.Value == "") 
            { return; }

            int RowNow = dataGridView1.CurrentRow.Index;//選択番号


            //現在のセルの値を表示
            memotxt = (string)dataGridView1.Rows[RowNow].Cells[3].Value;
            //現在のセルの値を表示
            memoName = (string)dataGridView1.Rows[RowNow].Cells[2].Value;

            txt.Txt_OUT_Edit(FoldaList+"\r\n\r\n" + memotxt);


            //メモ帳を起動する// ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();



            p.FileName = "notepad.exe";       // 起動するアプリケーション
            p.Arguments = FoldaList + "編集用テキスト.txt";            // 起動パラメータ
            p.UseShellExecute = false;                   // シェルを使用しない
            p.ErrorDialog = true;                        // 起動できなかった時にエラーダイアログを表示する
            p.ErrorDialogParentHandle = this.Handle;     // エラーダイアログを表示する親ハンドル(自フォーム)
            p.WorkingDirectory = Folda; // 多くは実行ファイルのディレクトリ
            p.CreateNoWindow = true;                     // コマンドプロンプトは非表示にする

            this.Size = size2;//縮小化

            // プロセスの起動
            System.Diagnostics.Process proc = System.Diagnostics.Process.Start(p);

            // プロセスが終了したときに Exited イベントを発生させる
            proc.EnableRaisingEvents = true;
            // Windows フォームのコンポーネントを設定して、コンポーネントが作成されているスレッドと
            // 同じスレッドで Exited イベントを処理するメソッドが呼び出されるようにする
            proc.SynchronizingObject = this;

            // プロセス終了時に呼び出される Exited イベントハンドラの設定
            proc.Exited += new EventHandler(Process_Exited);
        }


        #region 自分で定義したイベントその１
        // プロセスの終了を捕捉する Exited イベントハンドラ
        private void Process_Exited(object sender, EventArgs e)
        {
            txt.Txt_OUTEditEnd(FoldaList + "編集用テキスト");

            this.Size = size1;//サイズをもどす
            System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
            //System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
        }
        #endregion

        private void Text1_DoubleClick(object sender, EventArgs e)
        {
            txt.Txt_OUT_Edit(Text1.Text);//展開したサンプルを削除する

            //メモ帳を起動する// ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();



            p.FileName = "notepad.exe";       // 起動するアプリケーション
            p.Arguments = FoldaList + "編集用テキスト.txt";            // 起動パラメータ
            p.UseShellExecute = false;                   // シェルを使用しない
            p.ErrorDialog = true;                        // 起動できなかった時にエラーダイアログを表示する
            p.ErrorDialogParentHandle = this.Handle;     // エラーダイアログを表示する親ハンドル(自フォーム)
            p.WorkingDirectory = Folda; // 多くは実行ファイルのディレクトリ
            p.CreateNoWindow = true;                     // コマンドプロンプトは非表示にする

            // プロセスの起動
            System.Diagnostics.Process proc = System.Diagnostics.Process.Start(p);

            // プロセスが終了したときに Exited イベントを発生させる
            proc.EnableRaisingEvents = true;
            // Windows フォームのコンポーネントを設定して、コンポーネントが作成されているスレッドと
            // 同じスレッドで Exited イベントを処理するメソッドが呼び出されるようにする
            proc.SynchronizingObject = this;

            // プロセス終了時に呼び出される Exited イベントハンドラの設定
            proc.Exited += new EventHandler(Process1_Exited);
        }


        #region 自分で定義したイベントその２
        // プロセスの終了を捕捉する Exited イベントハンドラ
        private void Process1_Exited(object sender, EventArgs e)
        {
            txt.Txt_EditEnd(ref Text1);

            this.Size = size1;//サイズをもどす
            System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
            //System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
        }
        #endregion

        private void filePointbox00_SelectedIndexChanged(object sender, EventArgs e)
        {
            File_index_Number = filePointbox00.SelectedIndex;
        }



        private void テーブルの複製ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (comboBox0.Text == "")
            {
                MessageBox.Show("Accessのパスが指定されていません");
                return;
            }
            TableCopy(comboBox1, comboBox0);
        }

        #region テーブルの複製！！

        public void TableCopy(System.Windows.Forms.ComboBox comb1, System.Windows.Forms.ComboBox comb0)
        {
            string PrevTableName = comb1.Text;

            string s = Microsoft.VisualBasic.Interaction.InputBox(comb1.Text + " 　テーブルをコピーして作成するテーブル名");//インプットボックス使用
            if (txt.strListMacth(s, comb1)||txt.NG_Word_Chack(s)==false)  //一致する文字列を探す※既に登録されていないか判定
            {


                MessageBox.Show("すでに同じ名前があります");
                return;
            }
            else
            {
                DialogResult dr = MessageBox.Show("テーブルをコピーします。\n本当によろしいですか？\n" + "テーブル名：" + PrevTableName + "\r\nコピーする新規テーブル：" + s, "確認", MessageBoxButtons.OKCancel);
                if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    MessageBox.Show("Cancelを押しました。");
                    return;
                }
                else//Cancel以外の動作
                {

                    Program.SelectAccessPath = comb0.Text;
                    Program.SelectTablename = comb1.Text;

                    TableNameChengeSQL(Program.SelectAccessPath, Program.SelectTablename, s);

                    TableList();
                    comb1.SelectedIndex = txt.match(comb1, s);
                    Program.SelectTablename = comb1.Text;
                    LookTable(Program.SelectAccessPath, comb1);
                }
            }
        }
        #endregion




        #region サンプルテキストを複数展開
        private void MultiWorkBtn_Click(object sender, EventArgs e)
        {
                FolderBrowserDialog fbDialog = new FolderBrowserDialog();

                // ダイアログの説明文を指定する
                fbDialog.Description = "サンプルテキストを展開する場所を指定";

                switch (File_index_Number) {
                    case 0:
                        // デフォルトのフォルダを指定する
                        fbDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";//ダイアログを開いた最初の場所

                        break;
                    case 1:
                        // デフォルトのフォルダを指定する
                        fbDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";//ダイアログを開いた最初の場所

                        break;

                    case 2:
                        // デフォルトのフォルダを指定する
                        fbDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";//ダイアログを開いた最初の場所

                        break;
                    case 3:
                        // デフォルトのフォルダを指定する
                        fbDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";//ダイアログを開いた最初の場所

                        break;



                    default:
                        // デフォルトのフォルダを指定する
                        fbDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";//ダイアログを開いた最初の場所

                        break;
                }

                // 「新しいフォルダーの作成する」ボタンを表示する
                fbDialog.ShowNewFolderButton = true;

                //フォルダを選択するダイアログを表示する
                if (fbDialog.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show(fbDialog.SelectedPath);

                    string strp = fbDialog.SelectedPath;



                    GlobalPath = strp;

                    MultiWorking = true;

                }
                else
                {
                    Console.WriteLine("キャンセルされました");
                }

                // オブジェクトを破棄する
                fbDialog.Dispose();
        }
        #endregion



        #region テキストファイルの一斉読み込み
        private void テキストファイルの一斉読み込みToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DefaultNumber[1] = filePointbox00.SelectedIndex;

            Program.SelectAccessPath = comboBox0.Text;
            Program.SelectAccessname = comboBox00.Text;
            Program.SelectTablename = comboBox1.Text;

            ForcefullyReadOpen();
        }//一斉読み込み用のフォームへ移動する
        #endregion


        #region 項目行のテーブルを移動


        private void button10_Click(object sender, EventArgs e)
        {
            //////-----------------------------
            if (LookTableKEY == false)
            {
                MessageBox.Show("テーブルを選択してください。");
                return;
            }

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("データがありません");
                return;
            }
            this.Size = size2;//縮小化


            List<int> nRow = new List<int>();

            //選択されているすべての項目インデックス番号を取得する
            //dataGridView1.CurrentCell.RowIndex;

            foreach (DataGridViewRow d in dataGridView1.SelectedRows)//選択された行のみ
            {
                nRow.Add(d.Index);
            }
            string moji = "";
                for(int i=0;i<nRow.Count;i++) 
                {
                moji += "ID番号："+ dataGridView1.Rows[nRow[i]].Cells[0].Value.ToString()+"\n";
                 }

            DialogResult dr = MessageBox.Show("以下の項目を移動させます\n" +moji, "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
               // int RowNow = dataGridView1.CurrentRow.Index;//選択番号

                //// ヘッダ（項目名）
                //updFrm3[0, 0] = dataGridView1.Columns[0].HeaderText;//IDの番号
                //updFrm3[0, 1] = dataGridView1.Columns[1].HeaderText; ;//言語
                //updFrm3[0, 2] = dataGridView1.Columns[2].HeaderText;//名前
                //updFrm3[0, 3] = dataGridView1.Columns[3].HeaderText;//サンプル
                //updFrm3[0, 4] = dataGridView1.Columns[4].HeaderText;//ポイント

                for (int i=0;i<nRow.Count;i++) 
                {
                    //削除クエリに必要
                    Program.MulWork.nextWorks.Add(new FormNextWork());

                    Program.MulWork.nextWorks[i].PrevAccessPath = comboBox0.Text;
                    Program.MulWork.nextWorks[i].PrevAccessName = comboBox00.Text;
                    Program.MulWork.nextWorks[i].PrevTableName = comboBox1.Text;

                    //追加に必要
                    Program.MulWork.nextWorks[i].TableRowMenber.Add("ID", dataGridView1.Rows[nRow[i]].Cells[0].Value.ToString());//IDの番号
                    Program.MulWork.nextWorks[i].TableRowMenber.Add("言語", dataGridView1.Rows[nRow[i]].Cells[1].Value.ToString());//使用言語
                    Program.MulWork.nextWorks[i].TableRowMenber.Add("名前", dataGridView1.Rows[nRow[i]].Cells[2].Value.ToString());//名前
                    Program.MulWork.nextWorks[i].TableRowMenber.Add("サンプル", dataGridView1.Rows[nRow[i]].Cells[3].Value.ToString());//サンプル
                    Program.MulWork.nextWorks[i].TableRowMenber.Add("ポイント", dataGridView1.Rows[nRow[i]].Cells[4].Value.ToString());//ポイント
                }

    

                Program.FRow_Movecntrl = new FormRowMove();
                Program.FRow_Movecntrl.Show();
                Program.MainF.Hide();
            }
            else
            {
                MessageBox.Show("キャンセルしました");
                return;
            }
        }


        #endregion

        #region -------------【  ドラッグアンドドロップ  】---------------------------------------//

        private void Text1_DragEnter(object sender, DragEventArgs e) {
            // ドラッグ中のファイルやディレクトリの取得
            string[] sFileName = (string[])e.Data.GetData(System.Windows.Forms.DataFormats.FileDrop);

            //ファイルがドラッグされている場合、
            if (e.Data.GetDataPresent(System.Windows.Forms.DataFormats.FileDrop)) {
                // 配列分ループ
                foreach (string sTemp in sFileName) {
                    // ファイルパスかチェック
                    if (File.Exists(sTemp) == false) {
                        // ファイルパス以外なので何もしない
                        return;
                    } else {
                        break;
                    }
                }

                // カーソルを[+]へ変更する
                // ここでEffectを変更しないと、以降のイベント（Drop）は発生しない
                e.Effect = System.Windows.Forms.DragDropEffects.Copy;
            }
        }

        private void Text1_DragDrop(object sender, DragEventArgs e) {
            //ドロップされたファイルの一覧を取得
            string[] sFileName = (string[])e.Data.GetData(System.Windows.Forms.DataFormats.FileDrop, false);

            if (sFileName.Length <= 0) {
                return;
            }

            // ドロップ先がTextBoxであるかチェック
            System.Windows.Forms.TextBox TargetTextBox = sender as System.Windows.Forms.TextBox;

            if (TargetTextBox == null) {
                // TextBox以外のためイベントを何もせずイベントを抜ける。
                return;
            }

            // 現状のTextBox内のデータを削除
            TargetTextBox.Text = "";

            // TextBoxドラックされた文字列を設定
            TargetTextBox.Text = sFileName[0]; // 配列の先頭文字列を設定

            string txt = Text1.Lines[0];

           insertForm.txt.Reading(ref Text1, txt);

        }
        //------------    【   末尾  】     ----------------//
        #endregion


        private void btnSup_Click(object sender, EventArgs e) {
       

            this.Size = size2;//縮小化
            string openPath= Supple_Path(comboBox0.Text,comboBox1.Text);

            MessageBox.Show(openPath);
            Exe_Path.Explorer_Open(openPath);
        }

        public string Supple_Path(string AccessPath,string TableName) 
        {
          
            string a = Path_Cut.FileName(AccessPath);//Accessファイルへのパスからファイル名（拡張子なし）を取り出す

            return Exe_Path.exe_Path() + File_supple + "\\" + a + "\\" + TableName;
        }

        #region -------------【  検索機能  】---------------------------------------//
        private void btnSearch_Click(object sender, EventArgs e)
        {
            this.Size = size2;//縮小化

            if (Search_Method((int)Search_num.Name,ref dataGridView1,ref comboBox00,ref comboBox0,ref comboBox1)) {

                ReversePullForm f = new ReversePullForm();
                f.Show();
                f.textBox1.Text = KeyWord;
                f.dataGridView1 = SearchResultView;
                label1.Text = comboBox00.Text;
                this.Hide();
            }

        }//---------
    

        private void btnSearch2_Click(object sender, EventArgs e) {

            this.Size = size2;//縮小化

            if (Search_Method((int)Search_num.Sample, ref dataGridView1, ref comboBox00, ref comboBox0, ref comboBox1)) {

                ReversePullForm f = new ReversePullForm();
                f.Show();
                f.textBox1.Text = KeyWord;
                f.dataGridView1 = SearchResultView;
                label1.Text = comboBox00.Text;
                this.Hide();
            }

        }//--------------------

        public static bool Search_Method(int RETU,ref DataGridView data,ref System.Windows.Forms.ComboBox comboBox00,ref System.Windows.Forms.ComboBox comboBox0,ref System.Windows.Forms.ComboBox comboBox1) 
        {
            KeyWord = Microsoft.VisualBasic.Interaction.InputBox("検索キーワード");//インプットボックス使用

            if (KeyWord == "") {
                return false;
            }

            SearchResultView = new DataGridView();

            int num = data.Rows.Count;//行数
            int Row_num = 0;

            SearchResultView.ColumnCount = 5;



        for (int i = 0; i < num; i++) 
        {
            string Base = data.Rows[i].Cells[RETU].Value.ToString();//サンプルコードから検索

            if (StringSearch.IsSeach(Base, KeyWord)) {
                SearchResultView.ColumnCount = 5;
                SearchResultView.Rows.Add();
                SearchResultView.Rows[Row_num].Cells[0].Value = data.Rows[i].Cells[0].Value.ToString();
                SearchResultView.Rows[Row_num].Cells[1].Value = data.Rows[i].Cells[1].Value.ToString();
                SearchResultView.Rows[Row_num].Cells[2].Value = data.Rows[i].Cells[2].Value.ToString();
                SearchResultView.Rows[Row_num].Cells[3].Value = data.Rows[i].Cells[3].Value.ToString();
                SearchResultView.Rows[Row_num].Cells[4].Value = data.Rows[i].Cells[4].Value.ToString();

                Row_num++;
            }

        }

            if (Row_num <= 0) {
                MessageBox.Show(Row_num + " 件 : " + "見つかりませんでした");
                return false;
            } 
            else 
            {
                MessageBox.Show(Row_num + " 件 : " + "ヒットしました！");

                Program.SelectAccessname = comboBox00.Text;


                Program.SELECT_INDEX = comboBox1.SelectedIndex;
                Program.SelectTableList.Clear();

                Console.WriteLine(comboBox1.Text + " : " + comboBox1.Items.Count);

                for (int i = 0; i < comboBox1.Items.Count; i++) {
                    Program.SelectTableList.Add(comboBox1.Items[i].ToString());
                    MessageBox.Show(comboBox1.Items[i].ToString());
                }

            }


            for (int j = 0; j < 5; j++) 
            {
                SearchResultView.Columns[j].HeaderText = data.Columns[j].HeaderText;
            }

            return true;
        }

        enum Search_num 
       {
            ID=0,
            Lang=1,
            Name=2,
            Sample=3,
            Point=4
        }

        //------------    【   末尾  】     ----------------//
        #endregion


        private void tabControl1_TabIndexChanged(object sender, EventArgs e) {
           if(tabControl1.SelectedIndex == 1) 
           {
                textBox1.Focus();//フォーカスを当てる
            }
        }

        private void button4_Click_1(object sender, EventArgs e) {
            textBox1.Text = "";
        }

        private void btnEdit_Java_Click(object sender, EventArgs e) {
            //クリップボードに文字列をコピーする
            Clipboard.SetText("		System.out.println(\" \\n//-------------------------------//\\n\");");
        }

        private void btnEdit_VBA_Click(object sender, EventArgs e) {
            //クリップボードに文字列をコピーする
            Clipboard.SetText(@"'-------------------------------//");
        }

        private void btnEdit_Java2_Click(object sender, EventArgs e) {
            //クリップボードに文字列をコピーする
            Clipboard.SetText("//-------------------------------//");
        }


        private void excelを開くToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Office_Num office = Office_Num.EXCEL;
            string excellPth= MicrosoftApp.Get_OfficeAppPath(office);
            //Excel起動
            System.Diagnostics.Process p = System.Diagnostics.Process.Start(excellPth);//("excel.exe");

            //アイドル状態まで待機
            p.WaitForInputIdle();
        }

        private void excelで開くToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                // MessageBox.Show("データがありません");
                return;
            }
            int RowNow = dataGridView1.CurrentRow.Index;//選択番号

            List<string> Exmoji = new List<string>();

            //現在のセルの値を表示
            Exmoji=Str_Line.String_Add_Line( (string)dataGridView1.Rows[RowNow].Cells[3].Value);

            ExcelManager.Excel_Write(FoldaList + "編集用.xlsx",Exmoji);
            ExcelManager.OPEN_Excell(FoldaList + "編集用.xlsx");
        }

        private void FinalReleaseComObject(object o)
        {
            if (o != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
        }

        private void テーブルを展開ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int RowNow = dataGridView1.CurrentRow.Index;//選択番号

            //現在のセルの値を表示
            memotxt = (string)dataGridView1.Rows[RowNow].Cells[3].Value;

            //現在のセルの値を表示
            string name = (string)dataGridView1.Rows[RowNow].Cells[2].Value;

            txt.Txt_Multiple_Edit(GlobalPath, @"\編集用 " + C + name, memotxt);
            //MessageBox.Show(GlobalPath + @"\編集用  0  " + name + ".txt");
            //メモ帳起動
            System.Diagnostics.Process p = System.Diagnostics.Process.Start("notepad.exe", GlobalPath + @"\編集用 " + C + name + ".txt");

            C++;

            this.Size = size2;//縮小化
        }


        /* ------------------------------------------------------------------------------------------------ */


        #region 自分で定義したイベントえくせる
        // プロセスの終了を捕捉する Exited イベントハンドラ
        private void ProcessExcell_Exited(object sender, EventArgs e)
        {
            this.Size = size1;//サイズをもどす
            System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
            //System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
        }
        #endregion
   
        
        #region　mainFormを開く
        private void MainTool_exe_Click(object sender, EventArgs e) {


            //実行ファイルパス↓            
            string path = System.Windows.Forms.Application.ExecutablePath;

            string folderPath1 = Path.GetDirectoryName(path);

            string folderPath2 = Path.GetDirectoryName(folderPath1);

            
            string Mainstr = folderPath2 + @"\MainTools.exe";
            MessageBox.Show(Mainstr);

                //プログラムサンプルを起動
            System.Diagnostics.Process p = System.Diagnostics.Process.Start(Mainstr);

                //アイドル状態まで待機
                p.WaitForInputIdle();

                // this.Size = size2;//縮小化
                //自分自身のフォームを最小化
                this.WindowState = FormWindowState.Minimized;
         
        }
        #endregion
    }

}
