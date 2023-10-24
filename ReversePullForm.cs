using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace insertForm
{
    public partial class ReversePullForm : Form
    {
        public ReversePullForm()
        {
            InitializeComponent();
        }

        string memotxt;
        string memoName;
        int Row;
        int Colum;

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            this.dataGridView1.Font = new Font("Arial", (float)numericUpDown1.Value);
        }

        private void ReversePullForm_Load(object sender, EventArgs e)
        {
            //DataGridView1にユーザーが新しい行を追加できないようにする
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ColumnCount = 5;//dataSourceから値を入れないのであれば必要
            //dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
           
            //セルを選択すると行全体が選択されるようにする【行選択させるのに　すごく大事！！】
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            // dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            
            //DataGridView1の列の幅をユーザーが変更できないようにする
            dataGridView1.AllowUserToResizeColumns = true;

            //DataGridView1の行の高さをユーザーが変更できないようにする
            dataGridView1.AllowUserToResizeRows = true;


            for (int j = 0; j < 5; j++)
            {
                dataGridView1.Columns[j].HeaderText = Form1.SearchResultView.Columns[j].HeaderText;
            }

            int num = Form1.SearchResultView.Rows.Count;//行数
            //int Row_num = 0;

            for (int i = 0; i < num; i++)
            { 
                
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value=Form1.SearchResultView.Rows[i].Cells[0].Value;
                    dataGridView1.Rows[i].Cells[1].Value = Form1.SearchResultView.Rows[i].Cells[1].Value;
                    dataGridView1.Rows[i].Cells[2].Value = Form1.SearchResultView.Rows[i].Cells[2].Value;
                    dataGridView1.Rows[i].Cells[3].Value = Form1.SearchResultView.Rows[i].Cells[3].Value;
                    dataGridView1.Rows[i].Cells[4].Value = Form1.SearchResultView.Rows[i].Cells[4].Value;
            }



            //DataGridView1でセル、行、列が複数選択されないようにする【行選択させるのに大事！！】
            //dataGridView1.MultiSelect = false;

            dataGridView1.RowTemplate.Height = 100;
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 100;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 500;
            dataGridView1.Columns[4].Width = 500;

            //(0, 0)を現在のセルにする
            dataGridView1.CurrentCell = dataGridView1[0, 0];

            comboBox1.Items.AddRange(Program.SelectTableList.ToArray());
            comboBox1.SelectedIndex = Program.SELECT_INDEX;

            label2.Text = Program.SelectAccessname;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Program.MainF.Show();

            this.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            MessageBox.Show("Cell[" + Row + "," + Colum + "]");


            if (dataGridView1.RowCount == 0)
            {
                        MessageBox.Show("データがありません");
                        return;
            }
                  

             //memotxt = (string)dataGridView1.Rows[Row].Cells[Colum].Value;



            //現在のセルの値を表示
            memotxt = (string)dataGridView1.Rows[Row].Cells[3].Value;
            //現在のセルの値を表示
            memoName = (string)dataGridView1.Rows[Row].Cells[1].Value;



                    txt.Txt_OUT_Edit(Form1.FoldaList + memotxt);


                    //メモ帳を起動する// ProcessStartInfo の新しいインスタンスを生成する
                    System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();



                    p.FileName = "notepad.exe";       // 起動するアプリケーション
                    p.Arguments =Form1. FoldaList + "編集用テキスト.txt";            // 起動パラメータ
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
                txt.Txt_OUTEditEnd(Form1.FoldaList + "編集用テキスト");


                System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
                //System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
            }
        #endregion





        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show( $"クリックされた位置 {e.RowIndex}行目 {e.ColumnIndex}列目");
            Row = e.RowIndex;
            Colum = e.ColumnIndex;

            //// 行ヘッダー
            //if (e.RowIndex == -1)
            //{
            //    MessageBox.Show(s1 + Environment.NewLine + "行ヘッダーです。", "情報",
            //        MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}

            //// 列ヘッダー
            //if (e.ColumnIndex == -1)
            //{
            //    MessageBox.Show(s1 + Environment.NewLine + "列ヘッダーです。", "情報",
            //        MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}

            //string s2 = $"値は「{dataGridView1[e.ColumnIndex, e.RowIndex].Value}」です。";
            //MessageBox.Show(s1 + Environment.NewLine + s2, "情報",
            //    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
