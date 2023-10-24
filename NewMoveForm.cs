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
using System.Collections;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace insertForm
{
    public partial class NewMoveForm : Form {
        List<string> AcList;
        List<string> Out_AcList;
        List<string> ColumName = new List<string> { "言語", "名前", "サンプル" };
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        #region -------------【  Formクラスのコンストラクタ  】---------------------------------------//
        public NewMoveForm() {
            InitializeComponent();
        }


        public NewMoveForm(List<string> AcList) {
            InitializeComponent();
            this.AcList = AcList;
        }
        //------------    【   末尾  】     ----------------//
        #endregion


        public static string[] Movemojis;
        public static bool Ok = false;
        public static string memotxt;
        public static string memoname;
        private ContextMenuStrip contextMenuStrip1;
        private ContextMenuStrip contextMenuStrip2;

        #region ---------------【  Load　ロード！！！  】-----------------------------------------------
        private void FormRowMove_Load(object sender, EventArgs e) {
            dataGridView1.RowTemplate.Height = 40;
            dataGridView2.RowTemplate.Height = 40;

            // 行や幅をプログラムで設定するには、それらの自動設定が None になっている必要があるようです。
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            // ここでは DataGridView の全ての行に高さを設定しています。

            for (int row_index = 0; row_index < dataGridView1.Rows.Count; row_index++) {
                dataGridView1.Rows[row_index].Height = 10;
            }

            // 行や幅をプログラムで設定するには、それらの自動設定が None になっている必要があるようです。
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            // ここでは DataGridView の全ての行に高さを設定しています。

            for (int row_index = 0; row_index < dataGridView2.Rows.Count; row_index++) {
                dataGridView2.Rows[row_index].Height = 10;
            }

            //--------------------------------------（　＾ω＾）・・・
            string path = Application.ExecutablePath;
            string folderPath1 = Path.GetDirectoryName(path);
            string FolderPath = folderPath1 + @"\library\";//登録情報のテキストファイルへのパス

            //--------------------------------------（　＾ω＾）・・・
            //----------------------------------------------//
            IEnumerable<string> files = System.IO.Directory.EnumerateFiles(folderPath1, "*.accdb", System.IO.SearchOption.AllDirectories);

            //ファイルを列挙する
            foreach (string f in files) {
                listBox0.Items.Add(f);

                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                // string filePath = Path.GetFileNameWithoutExtension(f);

                //拡張子あり
                string filePath = System.IO.Path.GetFileName(f);


                listBox00.Items.Add(filePath);
            }
            //--------------------------------------------------//

            listBox0.SelectedIndex = txt.match0(listBox0, Program.MulWork.nextWorks[0].PrevAccessPath);
            listBox00.SelectedIndex = listBox0.SelectedIndex;


            TableList(listBox0,ref listBox1);

            if (Program.MulWork.nextWorks[0].NextTableName != "")
            {
                listBox1.SelectedIndex = txt.match0(listBox1, Program.MulWork.nextWorks[0].NextTableName);
            }
                #region dataGridView1へ代入
            //-------------------------------------------------------//
            //DataGridView1にユーザーが新しい行を追加できないようにする
            dataGridView1.AllowUserToAddRows = false;

            dataGridView1.ColumnCount = 5;//dataSourceから値を入れないのであれば必要

            Movemojis = new string[5];

            //表示
            dataGridView1.Rows.Add();//列の生成

            dataGridView1.Columns[0].HeaderText = "ID";
            dataGridView1.Columns[1].HeaderText = "言語";
            dataGridView1.Columns[2].HeaderText = "名前";
            dataGridView1.Columns[3].HeaderText = "サンプル";
            dataGridView1.Columns[4].HeaderText = "ポイント";

            dataGridView2.ColumnCount = 5;//dataSourceから値を入れないのであれば必要

            dataGridView2.Columns[0].HeaderText = "ID";
            dataGridView2.Columns[1].HeaderText = "言語";
            dataGridView2.Columns[2].HeaderText = "名前";
            dataGridView2.Columns[3].HeaderText = "サンプル";
            dataGridView2.Columns[4].HeaderText = "ポイント";

            //for (int i=0;i<5;i++) 
            //{
            //    Movemojis[i] = dataGridView1.Rows[0].Cells[i].Value.ToString();
            //}

            //Console.Write(resultDt.Rows[i][j] + " ");


            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 20;
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].Width = 100;

            dataGridView2.Columns[0].Width = 50;
            dataGridView2.Columns[1].Width = 20;
            dataGridView2.Columns[2].Width = 150;
            dataGridView2.Columns[3].Width = 100;
            dataGridView2.Columns[4].Width = 100;
            //"Column1"列のセルのテキストを折り返して表示する
            //dataGridView2.Columns[3].DefaultCellStyle.WrapMode =DataGridViewTriState.True;

            //------------------------------------------------//
            #endregion

            contextMenuStrip1 = new ContextMenuStrip();
            contextMenuStrip2 = new ContextMenuStrip();
            OutAcName00.ContextMenuStrip = contextMenuStrip1;//Accessファイルをコピー
            OutTableList.ContextMenuStrip = contextMenuStrip2;//-------テーブルリストボックスにセット

            this.contextMenuStrip1.Items.Add("Accessファイルをコピー",Properties.Resources.HoEn0, this.ContextMenu_Method1);
            this.contextMenuStrip1.Items.Add("Accessファイルを照合", Properties.Resources.HoEn1, this.ContextMenu_Method2);
            this.contextMenuStrip2.Items.Add("テーブルの内容を全てコピー", Properties.Resources.OK0, this.ContextMenu_Method3);
            this.contextMenuStrip2.Items.Add("データの照合コピー", Properties.Resources.OK1, this.ContextMenu_Method4);


        }//------------【　ロードイベント末尾　】
        #endregion

         #region ------【　Accessをコピーする　】--------
        private void ContextMenu_Method1(object sender, EventArgs e)
        {
            string AcFolder = Exe_Path.exe_Path() + @"Library\";
            List<string> result = new List<string>();
            int Ok_count = 0;

            //選択されているすべての項目インデックス番号を取得する
            for (int i = 0; i < OutAcName00.SelectedIndices.Count; i++)
            {

                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                string filePath = Path.GetFileNameWithoutExtension(@"C:Samurai\samurai.txt");
                if (File_Move_Copy.MoveingCpy1(OutAcPth0.Items[OutAcName00.SelectedIndices[i]].ToString(), AcFolder,ref result)) 
                {
                    Ok_count++;
                } 
            }

            for (int i = 0; i < Ok_count; i++) 
            {
                listBox0.Items.Add(result[i]); //拡張子あり
                listBox00.Items.Add(System.IO.Path.GetFileName(result[i]));
            }

            MessageBox.Show("処理完了  (≧ω≦)");
            preview(AcFolder);
            Aclist();
        }
        #endregion

        private void Aclist()
        {
            //--------------------------------------（　＾ω＾）・・・
            string path = System.Windows.Forms.Application.ExecutablePath;
            string folderPath1 = Path.GetDirectoryName(path);
            string af = folderPath1 + @"\library\";//登録情報のテキストファイルへのパス
            List<string> AcPathList=new List<string>();
            List<string> AcNameList = new List<string>();
            //--------------------------------------（　＾ω＾）・・・
            listBox0.Items.Clear();
            listBox00.Items.Clear();
            //----------------------------------------------//
            IEnumerable<string> files = System.IO.Directory.EnumerateFiles(af, "*.accdb", System.IO.SearchOption.TopDirectoryOnly);

            //ファイルを列挙する
            foreach (string f in files) {
                AcPathList.Add(f);
                //comboBox0.Items.Add(f);

            }
            //--------------------------------------------------//

            //ファイルを列挙する
            foreach (string f in files) {
                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                string filePath = Path.GetFileNameWithoutExtension(f);
                AcNameList.Add(filePath);
                //comboBox00.Items.Add(filePath);
            }
            //--------------------------------------------------//

            listBox0.Items.AddRange(AcPathList.ToArray());
            listBox00.Items.AddRange(AcNameList.ToArray());

        }

        #region -------【 Accessリスト「ファイルの照合」 】
        private void ContextMenu_Method2(object sender, EventArgs e)
        {
            int Table_Copy_count = 0;

            if (listBox00.Items.Count <= 0) { MessageBox.Show("ファイルがありません"); return; }

            OutAcPth0.SelectionMode = SelectionMode.One;
            OutAcName00.SelectionMode = SelectionMode.One;

            string AcFolder = Exe_Path.exe_Path() + @"Library\";
            List<string> result = new List<string>();
            int Ok_count = 0;

            //選択されているすべての項目インデックス番号を取得する
            for (int i = 0; i < OutAcName00.Items.Count; i++)
            {
                bool Is_File_flg = false;
                for (int j = 0; j < listBox0.Items.Count; j++)
                {
                    if (OutAcName00.Items[i].ToString() == listBox0.Items[j].ToString())
                    {
                        Is_File_flg = true;
                        break;
                    }
                }
                if (Is_File_flg)
                {
                    OutAcName00.SelectedIndex = i;
                    listBox00.SelectedIndex = i;
                    TableAutoCopy(ref Table_Copy_count);
                }
                else
                {
                    if (File_Move_Copy.MoveingCpy2(OutAcPth0.Items[i].ToString(), AcFolder, ref result))
                    {
                        Ok_count++;
                    }
                }
            }

            for (int i = 0; i < Ok_count; i++)
            {
                listBox0.Items.Add(result[i]); //拡張子あり
                listBox00.Items.Add(System.IO.Path.GetFileName(result[i]));
            }

            MessageBox.Show("処理完了  (≧ω≦)\r\n" + "Accessファイルコピー" + Ok_count + "件\r\n"+ "テーブルコピー" + Table_Copy_count + "件\r\n");
            preview(AcFolder);
            Aclist();
        }
        #endregion

        #region テーブルをコピーする
        private void ContextMenu_Method3(object sender, EventArgs e)
        {
            string Out_Table;

            //選択されているすべての項目を取得する
            Console.WriteLine("選択されているすべての項目：");
            for (int i = 0; i < OutTableList.SelectedItems.Count; i++)
            {
               Out_Table= OutTableList.SelectedItems[i].ToString();

               // MessageBox.Show("\nAccess：　"+OutAcPth0.Text+"\nテーブル: " +Out_Table);

                bool Flg = false;
                while (!Flg)
                {
                    string NewTableName = Microsoft.VisualBasic.Interaction.InputBox("変更しない場合は未入力", "コピーするテーブル名: " + Out_Table);


                    if (NewTableName == "")
                    {
                        NewTableName = Out_Table;
                    }

                    if (txt.Is_match(listBox1, NewTableName) == false)
                    {
                        DialogResult dr = MessageBox.Show("コピーテーブル：　" + Out_Table +"\nコピーAccess："+OutAcPth0.Text +"\r\n保存先Access：　" + listBox0.Text + "\r\n作成テーブル名：　" + NewTableName+"\n\n ※処理をキャンセルして終了する場合はキャンセルを選択", "確認", MessageBoxButtons.YesNoCancel);
                        if (dr == System.Windows.Forms.DialogResult.Cancel)
                        {
                            MessageBox.Show("Cancelを押しました\n処理を終了します");
                            return;
                        }
                        else
                        {
                            Access_SQL_Copy.TableNameChengeSQL_Another_AcMove(OutAcPth0.Text,Out_Table, listBox0.Text,NewTableName);
                          //  Multi_Table_CLONE(NewTableName,Prev_Table,NewAcPth0.Text);
                            break;
                        }
                    }
                    else
                    {
                        MessageBox.Show("同じ名前のテーブルがあります\n入力のやり直し");
                    }
                }
            }

            MessageBox.Show("処理完了  (≧ω≦)");
            listBox0.SelectedIndex = listBox00.SelectedIndex;
            dataGridView1.Rows.Clear();
            TableList(listBox0, ref listBox1);

            if (listBox1.Items.Count > 0) {
                listBox1.SelectedIndex = 0;
            }
        }
        #endregion

        #region テーブル 「照合コピー」
        private void ContextMenu_Method4(object sender, EventArgs e)//-------同じAccessファイルのテーブルデータを照合
        {
            int Table_copy_count = 0;
            TableAutoCopy(ref Table_copy_count);
            MessageBox.Show("\r\n"+ "テーブルコピー" + Table_copy_count + "件\r\n");
        }
        #endregion


        #region テーブル 「照合コピー」
        public void TableAutoCopy(ref int tCopy_count) 
        {
            OutTableList.SelectionMode = SelectionMode.One;
            dataGridView2.MultiSelect = false;
            
            tCopy_count= 0;
            int row_copy_count = 0;

            for (int i = 0; i < OutTableList.Items.Count; i++)
            {
                OutTableList.SelectedIndex = i;
                string tKey = OutTableList.Text;

                if (Mach(tKey)) //テーブルが一致した場合は
                {
                    //--dataGridView2.MultiSelect = false;
                  //  MessageBox.Show("テーブルに一致発見" + tKey);

                    foreach (DataGridViewRow d in dataGridView2.Rows)
                    {
                        //MessageBox.Show(d.Cells[3].Value.ToString());
                        if (dataMach_Row(d))//同じテーブル名の行で一致するサンプルがあるか
                        {
                            Movemojis[0] = d.Cells[0].Value.ToString();
                            Movemojis[1] = d.Cells[1].Value.ToString();
                            Movemojis[2] = d.Cells[2].Value.ToString();
                            Movemojis[3] = d.Cells[3].Value.ToString();

                            string Ac = listBox0.Text;
                            NextTableMove(listBox1, Ac);

                            List<List<string>> seQu = new List<List<string>>();

                            Access_SQL_Copy.Select_Query2(listBox0.Text, listBox1.Text, ref seQu);
                            Select_View(ref dataGridView1, seQu);
                            row_copy_count++;
                        }//同じサンプルデータがあれば追加しない

                    }

                }
                else //存在しないテーブルがある場合、テーブルをそのままコピーする
                {
                   // MessageBox.Show("一致しないテーブルをコピー");
                    string Out_TableName = OutTableList.Text;
                    Access_SQL_Copy.TableNameChengeSQL_Another_AcMove(OutAcPth0.Text, Out_TableName, listBox0.Text, Out_TableName);
                }
            }
            if (row_copy_count > 0)
            {
                MessageBox.Show("Accessファイル内\r\n総 処理件数："+row_copy_count + "件");
            }
            OutTableList.SelectionMode = SelectionMode.MultiExtended;
            dataGridView2.MultiSelect = true;
        }


        private bool dataMach_Row(DataGridViewRow r ) 
        {
                for(int i = 0; i < dataGridView1.RowCount; i++) 
                {
              //  MessageBox.Show("【判定】 "+dataGridView1.Rows[i].Cells[3].Value.ToString() +"  ==  "+ r.Cells[3].Value.ToString());
                   if( dataGridView1.Rows[i].Cells[3].Value.ToString() == r.Cells[3].Value.ToString()) 
                   {
                        return false;
                   }
                }
                return true;
        }


        private bool Mach(string Key) 
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                listBox1.SelectedIndex = i;
                if (listBox1.Text == Key) 
                {
                    return true;
                }
            }
            return false;
        }


        #endregion




        #region -------------【  削除候補　  】---------------------------------------//
        //private void Multi_Table_CLONE(string NewTableName,string PrevTable,string AcPath) 
        //{

        //    List<List<string>> SelectResult = new List<List<string>>();

        //  //  Access_SQL_Copy.Select_Query2(AcPath,PrevTable,ref SelectResult);

        //    foreach (List<string> Row_s in SelectResult)
        //    {
        //        Movemojis[0] = Row_s[0];
        //        Movemojis[1] = Row_s[1];
        //        Movemojis[2] = Row_s[2];
        //        Movemojis[3] = Row_s[3];

        //        string Ac = listBox0.Text;
        //        NextTableMove_And_Create(listBox1, Ac, NewTableName);
        //        Ok = true;
        //    }

        //}
        //------------    【   末尾  】     ----------------//
        #endregion



        #region -----【  選択行を複製したテーブルへ移動！！  】
        //public void NextTableMove_And_Create(ListBox list1, string Ac,string NewTableName)
        //{
        //    INSERT(list1, Ac);
        //    // Deleteting();
        //}
        #endregion



        #region ----------【 選択行を指定したテーブルへ移動！！】
        public  void NextTableMove(ListBox list1, string Ac) {

            INSERT(list1, Ac,Movemojis);
           // Deleteting();
        }
        #endregion



        #region ------【　追加クエリ　】------------------------
        public void INSERT(ListBox list1, string Ac,string[] Values) {
            if (list1.Text == "") {
                MessageBox.Show("テーブルを選択してください。");
                return;
            }
            List<string> Qu_Val = new List<string> { { Values[1] }, { Values[2] }, { Values[3] } };
            Access_SQL_Copy.INSERT_Query(Ac, list1.Text, ColumName, Qu_Val);
        }

        #endregion

        #region 削除クエリ
        public static void Deleteting() {
            //---------------------【テーブル名　取得　最初】----------------------------------//

            string AcPath = CONNECT_STRING + Program.MulWork.nextWorks[0].PrevAccessPath;

            string Keymoji = Program.MulWork.nextWorks[0].TableRowMenber["ID"];

            string Tablename = Program.MulWork.nextWorks[0].PrevTableName;

            string sql = "DELETE FROM " + Tablename + " WHERE ID = @p1";//IDの値を＠ｐ１に格納する
            OleDbConnection conect = new OleDbConnection(AcPath);
            conect.Open();
            OleDbCommand command = new OleDbCommand(sql, conect);
            command.Parameters.AddWithValue("@p1", Keymoji);//主キーの値
            int count = command.ExecuteNonQuery();

            MessageBox.Show("Delete Record:通知" + count);
            conect.Close();
            conect.Dispose();

        }


        #endregion


        #region -------------【  コピー実行ボタン  】---------------------------------------//
        private void button1_Click(object sender, EventArgs e) {
            if (listBox1.Text == ""||dataGridView2.CurrentCell.Value.ToString()=="") {
                MessageBox.Show("移動先のテーブルまたはコピーデータを選択してください");
                return;
            }

            //foreach( DataGridViewCell d in dataGridView2.SelectedCells){

            foreach (DataGridViewRow d in dataGridView2.SelectedRows)
            {

                Movemojis[0] = d.Cells[0].Value.ToString();
                Movemojis[1] = d.Cells[1].Value.ToString();
                Movemojis[2] = d.Cells[2].Value.ToString();
                Movemojis[3] = d.Cells[3].Value.ToString();

                string Ac = listBox0.Text;
                NextTableMove(listBox1, Ac);
                Ok = true;
            }

            List<List<string>> seQu = new List<List<string>>();

            Access_SQL_Copy.Select_Query2(listBox0.Text, listBox1.Text, ref seQu);
            Select_View(ref dataGridView1, seQu);
            //this.Close();
        }
        //------------    【   末尾  】     ----------------//
        #endregion

        #region -------------【  クローズイベント  】---------------------------------------//
        private void FormRowMove_FormClosing(object sender, FormClosingEventArgs e) {
            Program.MainF.Show();

            if (Ok == true) {
                Program.MainF.comboBox0.SelectedIndex = txt.match(Program.MainF.comboBox0, listBox0.Text);
                Program.MainF.comboBox00.SelectedIndex = Program.MainF.comboBox0.SelectedIndex;
                ;
                Program.MainF.comboBox1.SelectedIndex = txt.match(Program.MainF.comboBox1, listBox1.Text);

                Program.MainF.TableList();
                Form1.LookTable(Program.MainF.comboBox0.Text, Program.MainF.comboBox1);
            }

            Program.MulWork.nextWorks.Clear();
        }
        //------------    【   末尾  】     ----------------//
        #endregion


        #region -------------【 フォントサイズ   】---------------------------------------//
        private void numericUpDown1_ValueChanged(object sender, EventArgs e) {
            this.dataGridView1.DefaultCellStyle.Font = new Font("MS UI Gothic", (float)numericUpDown1.Value);
        }
        //------------    【   末尾  】     ----------------//
        #endregion


        #region -------------【 セルをダブルクリック   】---------------------------------------//
        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e) {
            int Coul = dataGridView1.CurrentCell.ColumnIndex;

            if (dataGridView1.CurrentCell.Value==null) {
                return;
            }


            //現在のセルの値を表示
            memotxt = (string)dataGridView1.Rows[0].Cells[Coul].Value;
            //現在のセルの値を表示
            memoname = dataGridView1.Columns[Coul].HeaderText;

            txt.Txt_OUT_Edit(memotxt);


            //メモ帳を起動する// ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();



            p.FileName = "notepad.exe";       // 起動するアプリケーション
            p.Arguments = "編集用テキスト.txt";            // 起動パラメータ
            p.UseShellExecute = false;                   // シェルを使用しない
            p.ErrorDialog = true;                        // 起動できなかった時にエラーダイアログを表示する
            p.ErrorDialogParentHandle = this.Handle;     // エラーダイアログを表示する親ハンドル(自フォーム)
            p.WorkingDirectory = Form1.Folda; // 多くは実行ファイルのディレクトリ
            p.CreateNoWindow = true;                     // コマンドプロンプトは非表示にする

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
        //------------    【   末尾  】     ----------------//
        #endregion


        #region 自分で定義したイベントその１
        // プロセスの終了を捕捉する Exited イベントハンドラ
        private void Process_Exited(object sender, EventArgs e) {
            txt.Txt_OUTEditEnd("編集用テキスト");


            System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
            System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
        }





        #endregion


        #region -------------【 ドラッグアンドドロップ   】---------------------------------------//
        private void NewAcName00_DragEnter(object sender, DragEventArgs e) {
            // ドラッグ中のファイルやディレクトリの取得
            string[] sFileName = (string[])e.Data.GetData(System.Windows.Forms.DataFormats.FileDrop);

            //ファイルがドラッグされている場合、
            if (e.Data.GetDataPresent(System.Windows.Forms.DataFormats.FileDrop)) {
                Folder folder = new Folder();

                // 配列分ループ
                foreach (string sTemp in sFileName) {
                    // ファイルパスかチェック
                    if (File.Exists(sTemp)) {
                        // ファイルパス以外なので何もしない

                    } else if (folder.Folder_Exis(sTemp)) {

                    } else {
                        continue;
                    }
                    // カーソルを[+]へ変更する
                    // ここでEffectを変更しないと、以降のイベント（Drop）は発生しない
                    e.Effect = System.Windows.Forms.DragDropEffects.Copy;
                }
            }
        }


        private void NewAcName00_DragDrop(object sender, DragEventArgs e) {
            OutAcName00.Items.Clear();
            OutAcPth0.Items.Clear();

            //コントロール内にドロップされたとき実行される
            //ドロップされたすべてのファイル名を取得する（フルパス）
            string[] fileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            List<string> FileDrop = new List<string>();


            FileDrop.AddRange(fileName);

            Drag_Drop.listCount(FileDrop, OutAcPth0, OutAcName00);
        }

        //------------    【   末尾  】     ----------------//
        #endregion


        #region -------------【 リストボックスの選択行変更 】---------------------------------------//
        private void NewAcName00_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void NewTableList_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e) {
            List<List<string>> seQu = new List<List<string>>();

            Access_SQL_Copy.Select_Query2(listBox0.Text, listBox1.Text, ref seQu);
            Select_View(ref dataGridView1, seQu);
        }
        private void listBox00_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox0.SelectedIndex = listBox00.SelectedIndex;
            dataGridView1.Rows.Clear();
            TableList(listBox0,ref listBox1);

            if (listBox1.Items.Count > 0)
            {
                listBox1.SelectedIndex = 0;
            }
        }

        //------------    【   末尾  】     ----------------//
        #endregion


        #region -------------【 選択クエリとテーブルの複製  】---------------------------------------//

        #region -------------【 テーブル名一覧   】---------------------------------------//
        private void TableList( ListBox list0,ref ListBox list1)
        {
            bool Connect_Flg = false;
            //---------------------【テーブル名　取得　最初】----------------------------------//
            list1.Items.Clear();
            List<string> TableName = new List<string>();

            Access_SQL_Copy.TableList2(list0.Text, ref TableName);

            list1.Items.AddRange(TableName.ToArray());
            if (list1.Items.Count > 0)
            {
                Connect_Flg = true;
                list1.SelectedIndex = 0;
            }
            else { Connect_Flg = false; }
        }
        //------------    【   末尾  】     ----------------//
        #endregion


        public void Select_View(ref DataGridView data ,List<List<string>> result) {
            //---------------------【テーブル名　取得　最初】----------------------------------//

            data.ColumnCount = 5;

            data.AllowUserToAddRows = false;

            //--見るだけの設定でも新規行は表示されない
            data.ReadOnly = true;

            data.Rows.Clear();

            if (result.Count > 0) {

                int Col = result[0].Count - 1;//列数-1

                for (int i = 0; i < result.Count; i++) {
                    data.Rows.Add();

                    for (int j = 0; j <= Col; j++) {

                        data.Rows[i].Cells[j].Value = result[i][j];
                    }
                }
                data.ReadOnly = true;
            }
        }

        #region テーブルの複製！！

        public static void TableCopy(ListBox MoveTableList,string PrevTableName,string PrevAcPath)
        {
           
            string NewTableName = Microsoft.VisualBasic.Interaction.InputBox(PrevTableName + " 　テーブルをコピーして作成するテーブル名","コピーテーブルの名前設定");//インプットボックス使用
            if (txt.strListMacth(NewTableName, MoveTableList) || txt.NG_Word_Chack(NewTableName) == false)  //一致する文字列を探す※既に登録されていないか判定
            {
                MessageBox.Show("すでに同じ名前_またはエラーが起きるテーブル名があります");
                return;
            }
            else
            {
                DialogResult dr = MessageBox.Show("テーブルをコピーします。\n本当によろしいですか？\n" + "テーブル名：" + PrevTableName, "確認", MessageBoxButtons.OKCancel);
                if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    MessageBox.Show("Cancelを押しました。");
                    return;
                }
                else//Cancel以外の動作
                {
                   Access_SQL_Copy.AccessNameChengeSQL(PrevAcPath, PrevTableName, NewTableName);
                }
            }
        }
        #endregion


        //------------    【   末尾  】     ----------------//
        #endregion

        #region -------------【  エクスプローラ開く  】---------------------------------------//
        public static void preview(string Path)
        {
            //エクスプローラでフォルダ"C:\My Documents\My Pictures"を開く
            System.Diagnostics.Process.Start("EXPLORER.EXE", Path);
        }


        //------------    【   末尾  】     ----------------//
        #endregion

        private void OutAcName00_SelectedIndexChanged(object sender, EventArgs e)
        {
            OutAcPth0.SelectedIndex = OutAcName00.SelectedIndex;
            dataGridView2.Rows.Clear();
            //OutTableList.Items.Clear();
            TableList(OutAcPth0, ref OutTableList);

            if (OutTableList.Items.Count > 0)
            {
                OutTableList.SelectedIndex = 0;
            }
        }

        private void OutTableList_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<List<string>> seQu = new List<List<string>>();

            Access_SQL_Copy.Select_Query2(OutAcPth0.Text, OutTableList.Text, ref seQu);
            Select_View(ref dataGridView2, seQu);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int Table_Copy_count = 0;

            if (listBox00.Items.Count <= 0) { MessageBox.Show("ファイルがありません");return; }

            OutAcPth0.SelectionMode = SelectionMode.One;
            OutAcName00.SelectionMode = SelectionMode.One;

            string AcFolder = Exe_Path.exe_Path() + @"Library\";
            List<string> result = new List<string>();
            int Ok_count = 0;

            //選択されているすべての項目インデックス番号を取得する
            for (int i = 0; i < OutAcName00.Items.Count; i++)
            {
                bool Is_File_flg = false;
                for(int j = 0; j < listBox00.Items.Count; j++) 
                {
                    if (OutAcName00.Items[i].ToString() == listBox00.Items[j].ToString())
                    {
                        Is_File_flg = true;
                        break;
                    }
                }
                if (Is_File_flg) 
                {
                    OutAcName00.SelectedIndex = i;
                    listBox00.SelectedIndex = i;
          
                    TableAutoCopy(ref Table_Copy_count);
                }
                else
                {
                    if (File_Move_Copy.MoveingCpy2(OutAcPth0.Items[i].ToString(), AcFolder, ref result))
                    {
                        Ok_count++;
                    }
                }
            }

            for (int i = 0; i < Ok_count; i++)
            {
                listBox0.Items.Add(result[i]); //拡張子あり
                listBox00.Items.Add(System.IO.Path.GetFileName(result[i]));
            }

            MessageBox.Show("処理完了  (≧ω≦)\r\n"+"Accessファイルコピー"+Ok_count+"件数"+ "\r\n" + "テーブルコピー" + Table_Copy_count + "件\r\n");
            preview(AcFolder);
            Aclist();
        }
    }
}
//https://dobon.net/vb/dotnet/datagridview/selectedcells.html
//https://dobon.net/vb/dotnet/control/lbselectitem.html
//https://dobon.net/vb/dotnet/datagridview/columncontains.html
//https://learn.microsoft.com/ja-jp/dotnet/api/system.windows.forms.datagridview.multiselect?view=windowsdesktop-7.0 複数行の選択
//https://ishii-singpg.com/archives/1801  Formクラスをコピーして使用（Formとして認識させる修正方法）