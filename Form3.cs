using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace insertForm
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        public static string[] daSt;
        public static string SelectTxt;
        public static bool UpDaFlg;
        public static DataGridView dagrVi2;
        public static int[] RowCoum;

        private void button1_Click(object sender, EventArgs e)//更新実行ボタン
        {
            DialogResult dr = MessageBox.Show("本当によろしいですか？", "確認", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.OK)
            {

                //---------------------【テーブル名　取得　最初】----------------------------------//
                string AcPath = CONNECT_STRING + Program.SelectAccessPath;

                //現在のセルの行インデックスを表示
                int nRow = dataGridView2.CurrentCell.RowIndex;

                daSt[0] = dataGridView2.Rows[nRow].Cells[0].Value.ToString();//言語
                daSt[1] = dataGridView2.Rows[nRow].Cells[1].Value.ToString();//名前
                daSt[2] = dataGridView2.Rows[nRow].Cells[2].Value.ToString();//サンプル
                daSt[3] = dataGridView2.Rows[nRow].Cells[3].Value.ToString();//ポイント
                daSt[4] = dataGridView2.Rows[nRow].Cells[4].Value.ToString();//Where条件のIDの値

                //---------------------------------------------//
                if (daSt[3] == "")
                {
                    daSt[3] = "-";
                }


                //MessageBox.Show(daSt[0]);
                //MessageBox.Show(daSt[1]);
                //MessageBox.Show(daSt[2]);
                //MessageBox.Show(daSt[3]);
                //MessageBox.Show(daSt[4]);
                //MessageBox.Show(Form1.CONNECT_STRING);
                //MessageBox.Show(Form1.tablename);
                //--------------------------------------------//

                string sql = "UPDATE " + Form1.tablename + " SET 言語=@p1,名前=@p2,サンプル=@p3,ポイント=@p4 WHERE ID=@p5";//条件は[ID]列の値
                OleDbConnection conect = new OleDbConnection(AcPath);

                conect.Open();
                OleDbCommand command = new OleDbCommand(sql, conect);
                command.Parameters.AddWithValue("@p1", daSt[1]);//言語
                command.Parameters.AddWithValue("@p2", daSt[2]);//名前
                command.Parameters.AddWithValue("@p3", daSt[3]);//サンプル
                command.Parameters.AddWithValue("@p4", daSt[4]);//ポイント
                command.Parameters.AddWithValue("@p5", daSt[0]);// Where条件の値【 ID 】

                //int count = command.ExecuteNonQuery();
                command.ExecuteNonQuery();
                MessageBox.Show("更新完了：");

                conect.Close();
                conect.Dispose();

              
                Program.MainF.Show();
                Program.MainF.Size = Form1.size1;
                Form1.LookTable(Program.MainF.comboBox0.Text,Program.MainF.comboBox1);
                Program.F3cntrl.Close();
            }//確認メッセージでOKを押す
        }

        private void button2_Click(object sender, EventArgs e)
        {   
            Program.MainF.Show();
            Program.F3cntrl.Close();
            
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            //DataGridView1にユーザーが新しい行を追加できないようにする
            dataGridView2.AllowUserToAddRows = false;

            daSt = new string[5];

            
            dataGridView2.ColumnCount = 5;//dataSourceから値を入れないのであれば必要

            //表示

            dataGridView2.Rows.Add();//列の生成

            for (int j = 0; j < 5; j++)
            {
                dataGridView2.Columns[j].HeaderText = Form1.updFrm3[0, j];
                dataGridView2.Rows[0].Cells[j].Value = Form1.updFrm3[1, j];

                //Console.Write(resultDt.Rows[i][j] + " ");
            }


        
            dataGridView2.Columns[0].Width = 50;
            dataGridView2.Columns[1].Width = 100;
            dataGridView2.Columns[2].Width = 250;
            dataGridView2.Columns[3].Width = 250;
            dataGridView2.Columns[4].Width = 250;
            //"Column1"列のセルのテキストを折り返して表示する
            //dataGridView2.Columns[3].DefaultCellStyle.WrapMode =DataGridViewTriState.True;

            dagrVi2 = dataGridView2;
            RowCoum = new int[2];
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Form3 f = new Form3();

            this.dataGridView2.Font = new Font("Arial", (float)numericUpDown1.Value);
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

        private void button4_Click(object sender, EventArgs e)
        {
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Minimized;
        }

        private void button5_Click(object sender, EventArgs e) {
            if (dataGridView2.CurrentCell.ColumnIndex == 0) {
                MessageBox.Show("ID列は変更できません");
                return;
            }

            SelectTxt = dataGridView2.CurrentCell.Value.ToString();
            UpDaFlg = false;

            RowCoum[0] = dataGridView2.CurrentCell.RowIndex;
            RowCoum[1] = dataGridView2.CurrentCell.ColumnIndex;
            //MessageBox.Show("行：" + RowCoum[0] + "　列：" + RowCoum[1]);

            Program.FUpDatecntrl = new FormUpDate();
            FormUpShowHide();
        }

        public void FormUpShowHide() {
            Program.FUpDatecntrl.Show();
            Program.F3cntrl.Hide();
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e) {
            if (dataGridView2.CurrentCell.ColumnIndex == 0) {
                MessageBox.Show("ID列は変更できません");
                return;
            }

            SelectTxt = dataGridView2.CurrentCell.Value.ToString();
            UpDaFlg = false;

            RowCoum[0] = dataGridView2.CurrentCell.RowIndex;
            RowCoum[1] = dataGridView2.CurrentCell.ColumnIndex;
           // MessageBox.Show("行："+RowCoum[0]+"　列："+RowCoum[1]);

            Program.FUpDatecntrl = new FormUpDate();
            FormUpShowHide();
        }
    }
}
