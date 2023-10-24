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

namespace insertForm
{
    public partial class FormRowMove : Form
    {
        public FormRowMove()
        {
            InitializeComponent();
        }

        public List< List< string>> Movemojis;
        public bool Ok=false;
        public static string memotxt;
        public static string memoname;
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        #region Load　ロード！！！！！！
        private void FormRowMove_Load(object sender, EventArgs e) {

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ReadOnly = true;

            Movemojis = new List<List<string>>();
            dataGridView1.RowTemplate.Height = 40;
    
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
                string filePath = Path.GetFileNameWithoutExtension(f);
                listBox00.Items.Add(filePath);
            }
            //--------------------------------------------------//

            listBox0.SelectedIndex = txt.match0(listBox0, Program.MulWork.nextWorks[0].PrevAccessPath);
            listBox00.SelectedIndex = listBox0.SelectedIndex;

            TableListMove(listBox0.Text, listBox1);

            labelAc.Text = Program.MulWork.nextWorks[0].PrevAccessName;
            labelT.Text = Program.MulWork.nextWorks[0].NextTableName;

            #region dataGridView1へ代入
            //-------------------------------------------------------//
            //DataGridView1にユーザーが新しい行を追加できないようにする
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;

            dataGridView1.ColumnCount = 5;//dataSourceから値を入れないのであれば必要

                dataGridView1.Columns[0].HeaderText = "ID";
                dataGridView1.Columns[1].HeaderText = "言語";
                dataGridView1.Columns[2].HeaderText = "名前";
                dataGridView1.Columns[3].HeaderText = "サンプル";
                dataGridView1.Columns[4].HeaderText = "ポイント";

            for (int i = 0; i < Program.MulWork.nextWorks.Count; i++) {

                //表示
                dataGridView1.Rows.Add();//行の生成




                //追加に必要
                dataGridView1.Rows[i].Cells[0].Value = Program.MulWork.nextWorks[i].TableRowMenber["ID"];
                dataGridView1.Rows[i].Cells[1].Value = Program.MulWork.nextWorks[i].TableRowMenber["言語"];
                dataGridView1.Rows[i].Cells[2].Value = Program.MulWork.nextWorks[i].TableRowMenber["名前"];
                dataGridView1.Rows[i].Cells[3].Value = Program.MulWork.nextWorks[i].TableRowMenber["サンプル"];
                dataGridView1.Rows[i].Cells[4].Value = Program.MulWork.nextWorks[i].TableRowMenber["ポイント"];


            } 
                for (int a = 0; a < Program.MulWork.nextWorks.Count; a++) {
                    Movemojis.Add(new List<string>());
              
                    for (int b=0;b<4;b++) {
                        Movemojis[a].Add(dataGridView1.Rows[a].Cells[b].Value.ToString());
                    } 
                }
            //Console.Write(resultDt.Rows[i][j] + " ");

            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].Width = 100;
            //"Column1"列のセルのテキストを折り返して表示する
            //dataGridView2.Columns[3].DefaultCellStyle.WrapMode =DataGridViewTriState.True;

            //------------------------------------------------//
            #endregion

        }
        #endregion

        #region テーブル名一覧！！
        public static void TableListMove(string Acse, ListBox list1)
        {
            //---------------------【テーブル名　取得　最初】----------------------------------//
            

            string AcPath= CONNECT_STRING + Acse;

       

            OleDbConnection conect = new OleDbConnection(AcPath);
            conect.Open();

            string[] opt = { null, null, null, "Table" };   //ユーザー定義のテーブルのみ取得を指定
            DataTable metaTable = conect.GetSchema("Tables", opt);   //データベースよりテーブル情報を取得
            DataColumn col = metaTable.Columns["TABLE_NAME"];

            if (list1.Items.Count>0)
            {
                list1.Items.Clear();
            }


            foreach (DataRow row in metaTable.Rows)
            {//-------------
                string tblName = row[col].ToString();//テーブル名

                string sql = "SELECT * FROM " + tblName;//--SQL Select文

                list1.Items.Add(tblName);//--テーブル名

            }//------------------【foreachの末尾】------------------

            conect.Close();
            conect.Dispose();

        }//---------------テーブル名一覧
        #endregion

        private void listBox00_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox0.SelectedIndex = listBox00.SelectedIndex;

            TableListMove(listBox0.Text, listBox1);
        }




        #region 選択行を指定したテーブルへ移動！！
        public  void NextTableMove(ListBox list1,string Ac,int RowIndex)
        {
          
            INSERT(list1,Ac,RowIndex);
            Deleteting(RowIndex);

        }

        #endregion
       


        #region 追加クエリ
        public void INSERT(ListBox list1 ,string Ac,int RowIndex) { 
            if (list1.Text=="")
            {
                MessageBox.Show("テーブルを選択してください。");
                return;
            }



            //---------------------【テーブル名　取得　最初】----------------------------------//
            string AcPath = CONNECT_STRING +  Ac;


            string TableNa = list1.Text;

            OleDbConnection conect = new OleDbConnection(AcPath);
            conect.Open();

            string sql = "INSERT INTO " + TableNa + " (言語,名前,サンプル) "   //ポイント
                + "VALUES (@a,@b,@c)";

            OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand型変数

            MessageBox.Show(Movemojis[RowIndex][2]);
            command.Parameters.AddWithValue("@a", Movemojis[RowIndex][1]);
            command.Parameters.AddWithValue("@b", Movemojis[RowIndex][2]);
            command.Parameters.AddWithValue("@c", Movemojis[RowIndex][3]);
            //command.Parameters.AddWithValue("@d", text3.Text);


            int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

             MessageBox.Show("追加件数：" + count);


            conect.Close();
            conect.Dispose();//ガベージコレクション（メモリの開放）
        }

        #endregion

        #region 削除クエリ
        public  void Deleteting(int RowIndex)
        {
            //---------------------【テーブル名　取得　最初】----------------------------------//

            string AcPath = CONNECT_STRING + Program.MulWork.nextWorks[RowIndex].PrevAccessPath; 

            string Keymoji = Program.MulWork.nextWorks[RowIndex].TableRowMenber["ID"];
            
            string Tablename= Program.MulWork.nextWorks[RowIndex].PrevTableName;

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

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.Text=="")
            {
                MessageBox.Show("移動先のテーブルを選択してください");
                return;
            }

            if (listBox1.Text == Program.MulWork.nextWorks[0].PrevTableName) {
                MessageBox.Show("違うテーブルを選択してください");
                return;
            }

            string Ac = listBox0.Text;

            int R_index = dataGridView1.CurrentCell.RowIndex;

            NextTableMove(listBox1,Ac,R_index);
            dataGridView1.Rows.RemoveAt(R_index);
            Movemojis.RemoveAt(R_index);
            if (dataGridView1.RowCount<=0) 
            {
              Ok = true;
              this.Close();
            }

        }

        private void FormRowMove_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.MainF.Show();

            if (Ok==true)
            {
                Program.MainF.comboBox0.SelectedIndex = txt.match(Program.MainF.comboBox0, listBox0.Text);
                Program.MainF.comboBox00.SelectedIndex = Program.MainF.comboBox0.SelectedIndex;
                ;
                Program.MainF.comboBox00.SelectedIndex = txt.match(Program.MainF.comboBox00, listBox1.Text);

                Program.MainF.TableList();
                Form1.LookTable(Program.MainF.comboBox0.Text, Program.MainF.comboBox1);
                Program.MainF.Size = new Size(978, 786);
    }

            Program.MulWork.nextWorks.Clear();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            this.dataGridView1.DefaultCellStyle.Font = new Font("Arial", (float)numericUpDown1.Value);

        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int Coul = dataGridView1.CurrentCell.ColumnIndex;

            if (Coul==0||Coul==1||Coul==2)
            {
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

        #region 自分で定義したイベントその１
        // プロセスの終了を捕捉する Exited イベントハンドラ
        private void Process_Exited(object sender, EventArgs e)
        {
            txt.Txt_OUTEditEnd("編集用テキスト");


            System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
            System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
        }





        #endregion

        private void FormRowMove_FormClosed(object sender, FormClosedEventArgs e) {

        }
    }
}
