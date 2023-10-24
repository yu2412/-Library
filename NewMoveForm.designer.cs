
namespace insertForm
{
    partial class NewMoveForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.listBox0 = new System.Windows.Forms.ListBox();
            this.listBox00 = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.label6 = new System.Windows.Forms.Label();
            this.OutAcPth0 = new System.Windows.Forms.ListBox();
            this.OutAcName00 = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.OutTableList = new System.Windows.Forms.ListBox();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(9, 314);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("MS UI Gothic", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dataGridView1.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowTemplate.Height = 100;
            this.dataGridView1.Size = new System.Drawing.Size(659, 435);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentDoubleClick);
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Font = new System.Drawing.Font("ＭＳ ゴシック", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(8, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 42);
            this.label2.TabIndex = 3;
            this.label2.Text = "移動先";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("ＭＳ ゴシック", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(668, 422);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 86);
            this.label3.TabIndex = 4;
            this.label3.Text = "←";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // listBox1
            // 
            this.listBox1.Font = new System.Drawing.Font("MS UI Gothic", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 27;
            this.listBox1.Location = new System.Drawing.Point(8, 55);
            this.listBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(308, 247);
            this.listBox1.TabIndex = 5;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // listBox0
            // 
            this.listBox0.FormattingEnabled = true;
            this.listBox0.ItemHeight = 15;
            this.listBox0.Location = new System.Drawing.Point(377, 64);
            this.listBox0.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.listBox0.Name = "listBox0";
            this.listBox0.Size = new System.Drawing.Size(108, 199);
            this.listBox0.TabIndex = 6;
            // 
            // listBox00
            // 
            this.listBox00.Font = new System.Drawing.Font("MS UI Gothic", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.listBox00.FormattingEnabled = true;
            this.listBox00.ItemHeight = 27;
            this.listBox00.Location = new System.Drawing.Point(352, 55);
            this.listBox00.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.listBox00.Name = "listBox00";
            this.listBox00.Size = new System.Drawing.Size(316, 247);
            this.listBox00.TabIndex = 7;
            this.listBox00.SelectedIndexChanged += new System.EventHandler(this.listBox00_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.White;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Font = new System.Drawing.Font("ＭＳ ゴシック", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(349, 5);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(160, 42);
            this.label4.TabIndex = 8;
            this.label4.Text = "Access";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.White;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label5.Font = new System.Drawing.Font("ＭＳ Ｐ明朝", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(105, 10);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(141, 42);
            this.label5.TabIndex = 9;
            this.label5.Text = "テーブル名";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(756, 422);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(158, 73);
            this.button1.TabIndex = 14;
            this.button1.Text = "実行";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Font = new System.Drawing.Font("MS UI Gothic", 25.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.numericUpDown1.Location = new System.Drawing.Point(764, 5);
            this.numericUpDown1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            40,
            0,
            0,
            0});
            this.numericUpDown1.Minimum = new decimal(new int[] {
            9,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(95, 50);
            this.numericUpDown1.TabIndex = 15;
            this.numericUpDown1.Value = new decimal(new int[] {
            14,
            0,
            0,
            0});
            this.numericUpDown1.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(972, 308);
            this.dataGridView2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("MS UI Gothic", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dataGridView2.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView2.RowTemplate.Height = 100;
            this.dataGridView2.Size = new System.Drawing.Size(635, 418);
            this.dataGridView2.TabIndex = 16;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.White;
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label6.Font = new System.Drawing.Font("ＭＳ ゴシック", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(1349, 4);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(132, 42);
            this.label6.TabIndex = 18;
            this.label6.Text = "Access";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // OutAcPth0
            // 
            this.OutAcPth0.Font = new System.Drawing.Font("MS UI Gothic", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.OutAcPth0.FormattingEnabled = true;
            this.OutAcPth0.ItemHeight = 27;
            this.OutAcPth0.Location = new System.Drawing.Point(1394, 49);
            this.OutAcPth0.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OutAcPth0.Name = "OutAcPth0";
            this.OutAcPth0.Size = new System.Drawing.Size(35, 220);
            this.OutAcPth0.TabIndex = 17;
            // 
            // OutAcName00
            // 
            this.OutAcName00.AllowDrop = true;
            this.OutAcName00.Font = new System.Drawing.Font("MS UI Gothic", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.OutAcName00.FormattingEnabled = true;
            this.OutAcName00.ItemHeight = 27;
            this.OutAcName00.Location = new System.Drawing.Point(1341, 49);
            this.OutAcName00.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OutAcName00.Name = "OutAcName00";
            this.OutAcName00.Size = new System.Drawing.Size(266, 247);
            this.OutAcName00.TabIndex = 19;
            this.OutAcName00.SelectedIndexChanged += new System.EventHandler(this.OutAcName00_SelectedIndexChanged);
            this.OutAcName00.DragDrop += new System.Windows.Forms.DragEventHandler(this.NewAcName00_DragDrop);
            this.OutAcName00.DragEnter += new System.Windows.Forms.DragEventHandler(this.NewAcName00_DragEnter);
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.White;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label7.Font = new System.Drawing.Font("ＭＳ Ｐ明朝", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(973, 5);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(141, 42);
            this.label7.TabIndex = 21;
            this.label7.Text = "テーブル名";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // OutTableList
            // 
            this.OutTableList.Font = new System.Drawing.Font("MS UI Gothic", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.OutTableList.FormattingEnabled = true;
            this.OutTableList.ItemHeight = 27;
            this.OutTableList.Location = new System.Drawing.Point(973, 50);
            this.OutTableList.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OutTableList.Name = "OutTableList";
            this.OutTableList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.OutTableList.Size = new System.Drawing.Size(348, 247);
            this.OutTableList.TabIndex = 20;
            this.OutTableList.SelectedIndexChanged += new System.EventHandler(this.OutTableList_SelectedIndexChanged);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(728, 184);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(198, 118);
            this.button2.TabIndex = 22;
            this.button2.Text = "Autoコピー";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // NewMoveForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1645, 751);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.OutTableList);
            this.Controls.Add(this.OutAcName00);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.OutAcPth0);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.listBox00);
            this.Controls.Add(this.listBox0);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label3);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "NewMoveForm";
            this.Text = "データの受け渡し";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormRowMove_FormClosing);
            this.Load += new System.EventHandler(this.FormRowMove_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.ListBox listBox0;
        private System.Windows.Forms.ListBox listBox00;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox OutAcPth0;
        private System.Windows.Forms.ListBox OutAcName00;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ListBox OutTableList;
        private System.Windows.Forms.Button button2;
    }
}