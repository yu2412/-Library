using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace insertForm {
    public partial class FormUpDate : Form {

        public static string Folda;



        public FormUpDate() {
            InitializeComponent();
        }

        private void FormUpDate_Load(object sender, EventArgs e) {
            richTextBox1.Text = Form3.SelectTxt;

            string path = System.Windows.Forms.Application.ExecutablePath;

            string folderPath = System.IO.Path.GetDirectoryName(path);

            Folda = folderPath + @"\";
        }

        public void hideForms2() {
            Program.F3cntrl.Show();
            Program.MainF.Show();
           
        }

        private void 最小化ToolStripMenuItem_Click(object sender, EventArgs e) {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;

        }

        private void 最大化通常ToolStripMenuItem_Click(object sender, EventArgs e) {
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

        private void button2_Click(object sender, EventArgs e) {
            Form3.SelectTxt = "";
            Program.F3cntrl.Show();
            Program.FUpDatecntrl.Close();
        }

        private void button1_Click(object sender, EventArgs e) {
            Form3.SelectTxt = richTextBox1.Text;
            
            Form3.UpDaFlg = true;
            Program.FUpDatecntrl.Close();
        }

        private void FormUpDate_FormClosing(object sender, FormClosingEventArgs e) {
            Program.F3cntrl.Show();

            if (Form3.UpDaFlg==true) {
                Form3.dagrVi2.Rows[Form3.RowCoum[0]].Cells[Form3.RowCoum[1]].Value = Form3.SelectTxt;
                Form3.UpDaFlg = false;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            this.richTextBox1.Font = new Font("Arial", (float)numericUpDown1.Value);
        }

        private void richTextBox1_DoubleClick(object sender, EventArgs e)
        {

            //txt.Txt_OUT_Edit(richTextBox1.Text);

            ////メモ帳を起動する// ProcessStartInfo の新しいインスタンスを生成する
            //System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();



            //p.FileName = "notepad.exe";       // 起動するアプリケーション
            //p.Arguments = "編集用テキスト.txt";            // 起動パラメータ
            //p.UseShellExecute = false;                   // シェルを使用しない
            //p.ErrorDialog = true;                        // 起動できなかった時にエラーダイアログを表示する
            //p.ErrorDialogParentHandle = this.Handle;     // エラーダイアログを表示する親ハンドル(自フォーム)
            //p.WorkingDirectory = Folda; // 多くは実行ファイルのディレクトリ
            //p.CreateNoWindow = true;                     // コマンドプロンプトは非表示にする

            //// プロセスの起動
            //System.Diagnostics.Process proc = System.Diagnostics.Process.Start(p);

            //// プロセスが終了したときに Exited イベントを発生させる
            //proc.EnableRaisingEvents = true;
            //// Windows フォームのコンポーネントを設定して、コンポーネントが作成されているスレッドと
            //// 同じスレッドで Exited イベントを処理するメソッドが呼び出されるようにする
            //proc.SynchronizingObject = this;

            //// プロセス終了時に呼び出される Exited イベントハンドラの設定
            //proc.Exited += new EventHandler(Process_Exited);
        }

        // プロセスの終了を捕捉する Exited イベントハンドラ
        private void Process_Exited(object sender, EventArgs e)
        {
            txt.Txt_EditEnd(ref richTextBox1);


            System.Diagnostics.Process proc = (System.Diagnostics.Process)sender;
            System.Windows.Forms.MessageBox.Show("プロセスが終了しました。プロセスID：" + proc.Id.ToString());
        }


    }
}
