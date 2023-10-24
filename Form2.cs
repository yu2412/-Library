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
using System.Threading;

namespace insertForm {
    public partial class Form2 : Form {
        public Form2() {
            InitializeComponent();
        }

        public static string Folda;


        private void Form2_Load(object sender, EventArgs e) {

            // アイコンファイルのパス
            string p =Form1.FoldaList + @"usa2.ico";
            // パスを指定してアイコンのインスタンスを生成
            Icon icon = new Icon(p, 16, 16);
            // フォームのIconプロパティに設定
            this.Icon = icon;

            richTextBox1.Text = Form1.memotxt;

            Namelbl.Text = Form1.memoName;

            string path = System.Windows.Forms.Application.ExecutablePath;

            string folderPath = System.IO.Path.GetDirectoryName(path);

            Folda = folderPath + @"\";
        }

        private void 戻るToolStripMenuItem_Click(object sender, EventArgs e) {
            if (ForcefullyReadForm.WorktxtFlg==true) {
                ForcefullyReadForm.WorktxtFlg = false;

                Program.ForcefullyReadCntrl.Show();
                this.Close();

            } else {

                hideForms2();
            }
        }



        public void hideForms2() {
            Program.MainF.Show();
            Program.F2cntrl.Close();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e) {
            
                Form2 f = new Form2();

                this.richTextBox1.Font = new Font("Arial", (float)numericUpDown1.Value);
            
        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e) {
            //if (textBox2.Text == "" || textBox1.Text == "") {
            if (richTextBox1.Text == "") {
                MessageBox.Show("保存プログラムを選択してください");
                return;
            }

            SaveFileDialog dialog = new SaveFileDialog();

            //dialog.InitialDirectory = @"C:\Users\std08\Downloads\C#\";
            //dialog.InitialDirectory= @"C:\Users\higay\Downloads\";
            dialog.InitialDirectory = @"C:\";//Users\std08\Downloads\C#\";


            dialog.FileName =Namelbl.Text;

            dialog.Filter = "テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*";
            //dialog.Title = "ファイルを保存する";
            if (dialog.ShowDialog() == DialogResult.OK) {

                File.WriteAllText(dialog.FileName, richTextBox1.Text);



            }
        }

        private void 最大化通常ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void 最小化ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;
        }

        private void button1_Click(object sender, EventArgs e) {

            string S1 = richTextBox1.Text;

            txt.TxtStdWrite(Namelbl.Text+".txt",S1);

                    //メモ帳起動
                    System.Diagnostics.Process p = System.Diagnostics.Process.Start("notepad.exe",Namelbl.Text+".txt");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //Excel起動
            System.Diagnostics.Process p = System.Diagnostics.Process.Start("excel.exe");

            //アイドル状態まで待機
            p.WaitForInputIdle();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            txt.txtDelete(Namelbl.Text+".txt");
        }

        private void richTextBox1_DoubleClick(object sender, EventArgs e)
        {


        }

        private void richTextBox1_TextChanged(object sender, EventArgs e) {

        }
    }
}
