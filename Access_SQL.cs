using insertForm;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MyCreate
{


   public class Access_SQL_Copy
   {
        static string File_supple = "Library_Supple";//-------各テーブル用のフォルダ


        static  string[] Query_value = {"@a" , "@b" , "@c" , "@d" , "@e" , "@f" , "@g" , "@h" , "@i" , "@j" , "@k" , "@l" , "@m" , "@n" , "@o" , "@p" , "@q" , "@r" , "@s" , "@t" , "@u" , "@v" , "@w" , "@x" , "@y" , "@z" };


        
        const string CONNECT_STRING = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source= ";
        

        const string STRING_END = ";";
        //const string TEXT_PS = @"PS_List\ps.txt";//".env";
        const string TEXT_PS = @"PS_List\.en";

        static string Pstxt;

        static bool HomePs_Flg = false;
        static string HomePs = "";


        #region テーブル複製メソッド＿コピーを別名で作成

        public static void AccessNameChengeSQL(string Ac, string TaNa, string NextName)
        {

            //---------------------【テーブル名　取得　最初】----------------------------------//
            string AcPath = CONNECT_STRING + Ac;

            //MessageBox.Show(Form1.CONNECT_STRING);
            //MessageBox.Show(TaNa);
            //--------------------------------------------//

            string sql = "SELECT  * INTO " + NextName + " FROM " + TaNa;

            OleDbConnection conect_prev = new OleDbConnection(AcPath);
            conect_prev.Open();

            try
            {

                OleDbCommand command = new OleDbCommand(sql, conect_prev);  //OleDbCommand　命令を送る
                int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

                MessageBox.Show("テーブル複製：" + count);

                string a = Path_Cut.FileName(Ac);

                string path1 = Exe_Path.exe_Path() + File_supple + "\\" + a + "\\" + TaNa;
                string path2 = Exe_Path.exe_Path() + File_supple + "\\" + a + "\\" + NextName;

                Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(path1, path2);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            conect_prev.Close();
            conect_prev.Dispose();//ガベージコレクション（メモリの開放）
        }

        #endregion

        #region テーブル複製メソッド＿コピーを別名で作成

        public static void TableNameChengeSQL_Another_AcMove(string Out_Ac, string Out_TaNa, string NewAcPath,string NewTableName)
        {
            List<List<string>> result = new List<List<string>>();

            List<string> ColumName = new List<string> { "言語","名前","サンプル" };

            #region  -----------【  前のAccessからSELECTクエリで取得  】---------------------------------------//

            Select_Query2(Out_Ac, Out_TaNa, ref result);

            List<List<string>> Remake = new List<List<string>>();

            for(int i=0;i<result.Count;i++ ) 
            {
                Remake.Add(new List<string>());
              // MessageBox.Show(i+":"+result[i][2]);
                for (int j = 1; j <= 3; j++) 
                {
                   
                    Remake[i].Add(result[i][j]);
                }
            }

            //------------    【   末尾  】     ----------------//
            #endregion
            if (Create_Table_Query(NewAcPath, NewTableName))
            {
                int count = 0;
                foreach (List<string> Row_s in  Remake)
                {
                       if(INSERT_Query(NewAcPath, NewTableName, ColumName, Row_s)) 
                       {
                        count++;
                       }
                }
                MessageBox.Show("追加件数："+count+"件");
            }
            else
            {
                MessageBox.Show("エラーテーブルの作成に失敗");
            }
        }
        #endregion

        #region テーブルの新規作成

        public static bool Create_Table_Query(string AcPath, string newTableName)
        {
            bool re = false;

            string acpath = CONNECT_STRING + AcPath;

            if (File.Exists(AcPath) == false) { return false; }

            OleDbConnection conect = new OleDbConnection(acpath);

            conect.Open();

            try
            {

                string sql = "CREATE TABLE " + newTableName + " (ID AUTOINCREMENT PRIMARY KEY, 言語 TEXT ,名前 TEXT,サンプル TEXT, ポイント TEXT)";//オートナンバー型に主キーの設定（PRIMARY KEY）//textBox1.Textは新規作成のテーブル名

                OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand　命令を送る
                int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

               // MessageBox.Show("追加テーブル件数：" + count);

                re = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                re = false;
            }

            conect.Close();
            conect.Dispose();//ガベージコレクション（メモリの開放）
            return re;
        }

        #endregion


        #region テーブル名一覧表示　その２

        public static void TableList2(string exePath, ref List<string> list) {
            string acpath = CONNECT_STRING + exePath;

            //---------------------【テーブル名　取得　最初】----------------------------------//


            if (File.Exists(exePath) == false) {
               // MessageBox.Show("ファイルが存在しません");
                return;
            }


            OleDbConnection conect = new OleDbConnection(acpath);


            conect.Open();

            string[] opt = { null, null, null, "Table" };   //ユーザー定義のテーブルのみ取得を指定
            DataTable metaTable = conect.GetSchema("Tables", opt);   //データベースよりテーブル情報を取得
            DataColumn col = metaTable.Columns["TABLE_NAME"];


            foreach (DataRow row in metaTable.Rows) {//-------------
                string tblName = row[col].ToString();//テーブル名

                string sql = "SELECT * FROM " + tblName;//--SQL Select文

                list.Add(tblName);//--テーブル名

            }//------------------【foreachの末尾】------------------

            conect.Close();
            conect.Dispose();

            //--------------------【テーブル名　取得　末尾】---------------------------------//



        }//テーブル名を一覧取得メソッド

        #endregion



        #region テーブル名一覧表示

        public static void TableList(string exePath,ref ListBox list)
        {
            string acpath = CONNECT_STRING + exePath;

            //---------------------【テーブル名　取得　最初】----------------------------------//


            if (File.Exists(exePath)==false)
            {
                MessageBox.Show("ファイルが存在しません");
                return;
            }


            OleDbConnection conect = new OleDbConnection(acpath);


                conect.Open();

                string[] opt = { null, null, null, "Table" };   //ユーザー定義のテーブルのみ取得を指定
                DataTable metaTable = conect.GetSchema("Tables", opt);   //データベースよりテーブル情報を取得
                DataColumn col = metaTable.Columns["TABLE_NAME"];


                foreach (DataRow row in metaTable.Rows)
                {//-------------
                    string tblName = row[col].ToString();//テーブル名

                    string sql = "SELECT * FROM " + tblName;//--SQL Select文

                    list.Items.Add(tblName);//--テーブル名

                }//------------------【foreachの末尾】------------------

                conect.Close();
                conect.Dispose();

                //--------------------【テーブル名　取得　末尾】---------------------------------//



        }//テーブル名を一覧取得メソッド

        #endregion



        #region 選択クエリ　DataGridView
        public static void Select_Query(string AcPath,string TableName,ref DataGridView data)
        {
            string acpath = CONNECT_STRING + AcPath;

            if (File.Exists(AcPath)==false) { return ; }

                OleDbConnection conect = new OleDbConnection(acpath);

            //conect.Open();
            string sql = "SELECT * FROM " + TableName + " ORDER BY ID";//TestSampletableをID順に
            //textBox1.Text = sql;


            OleDbDataAdapter da = new OleDbDataAdapter(sql, conect);


            DataTable tbl = new DataTable();

            da.Fill(tbl);

            data.DataSource = tbl;
        }//選択クエリの末尾
        #endregion

        #region 選択クエリ その２
        public static void Select_Query2(string AcPath, string TableName, ref List< List<string>> Result)
        {
           
            string acpath = CONNECT_STRING  +AcPath;
           // MessageBox.Show(acpath);


            if (File.Exists(AcPath) == false) { return; }

            OleDbConnection conect = new OleDbConnection(acpath);

            //conect.Open();
            string sql = "SELECT * FROM " + TableName;
            //textBox1.Text = sql;


            OleDbDataAdapter da = new OleDbDataAdapter(sql, conect);


            DataTable tbl = new DataTable();

            da.Fill(tbl);

          
            //MessageBox.Show(tbl.Rows[0].ItemArray[0].ToString());

            for (int i = 0; i < tbl.Rows.Count; i++) 
            {
                List<string> Colum = new List<string>();
                Result.Add(Colum) ;

                for (int j = 0; j < tbl.Columns.Count; j++) 
                {
                   // MessageBox.Show(tbl.Rows[i].ItemArray[j].ToString());
                    Result[i].Add(tbl.Rows[i].ItemArray[j].ToString());
                }
            }

            
        }//選択クエリの末尾


        #endregion

        #region 文字列の組み立て
        public static  string String_Joint(List<string> Row_Name,string SQL_Type,string TableName) 
        {
            string Joint_Row_sum = " (";
            string Joint_str = "";

            Joint_str +=  Row_Name[0];

            for (int i=1;i<Row_Name.Count;i++) 
            {
                Joint_str += ","+Row_Name[i];
            }

            Joint_str += ") VALUES (@a";


            for (int i = 1; i < Row_Name.Count; i++)
            {
                Joint_str += "," + Query_value[i];
            }

            Joint_str += ")";

            string Result = SQL_Type + TableName + Joint_Row_sum+Joint_str;

            return Result;
        }


        #endregion

        #region 挿入クエリ

        public static bool INSERT_Query(string AcPath, string TableName,List<string> Coulum_Name ,List<string> Coulum_Value) 
        {
            bool re = false;

            string acpath = CONNECT_STRING + AcPath;

            if ((Coulum_Name.Count)!=Coulum_Value.Count) { MessageBox.Show("列名と値の数が違います");return false; }

            if (File.Exists(AcPath) == false) { return false; }

            OleDbConnection conect = new OleDbConnection(acpath);

            conect.Open();
        
            string sql = String_Joint(Coulum_Name, "INSERT INTO ", TableName);
           //  MessageBox.Show(sql);
            //string sql = "INSERT INTO " + TableName + "(URL,説明)" + "VALUES (@a,@b)";

            try
            {

                OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand型変数

                for (int i = 0; i < Coulum_Name.Count; i++)
                {
                   // MessageBox.Show(" command.Parameters.AddWithValue(" + Query_value[i]+" , "+ Coulum_Value[i]);
                    command.Parameters.AddWithValue(Query_value[i], Coulum_Value[i]);
                
                }

                int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

               // MessageBox.Show("追加件数：" + count);
                re = true;
            }
            catch (Exception e) 
            {
                MessageBox.Show(e.Message);
                Console.WriteLine(e.Message);
                re = false;
            }

            conect.Close();
            conect.Dispose();//ガベージコレクション（メモリの開放）
            return re;
        }

        #endregion



        #region 削除クエリ
        public static bool Delete_Query(string AcPath, string TableName, string Key)
        {
            bool re = false;
            string acpath = CONNECT_STRING + AcPath;

            if (File.Exists(AcPath) == false) { return false; }


            string ConditionsStr = Key;

            string sql = "DELETE FROM " + TableName + " WHERE ID = @p1";//IDの値を＠ｐ１に格納する
            OleDbConnection conect = new OleDbConnection(acpath);
            
            try
            {
                
                conect.Open();
                OleDbCommand command = new OleDbCommand(sql, conect);
                command.Parameters.AddWithValue("@p1", ConditionsStr);//主キーの値
                int count = command.ExecuteNonQuery();

                MessageBox.Show("Delete Record:通知" + count);
                re = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                re = false;
            }


            conect.Close();
            conect.Dispose();

            return re;
        }//選択クエリの末尾


        #endregion

        #region 文字列の組み立て
        public static string String_Joint_Update(List<string> Coulum_Name)
        {
            string Joint_Row = "";
        

            Joint_Row += Coulum_Name[0]+"="+Query_value[1];

            for (int i = 1; i < Coulum_Name.Count; i++)
            {
                Joint_Row += "," + Coulum_Name[i]+" = "+Query_value[i+1];
            }

            //"言語=@b,名前=@p2,サンプル=@p3,ポイント=@p4"

            string Result = Joint_Row;

            return Result;
        }


        #endregion




        #region 更新Query
        public static bool UpDate_Query(string AcPath, string TableName,string id_Name,string id_Value, List<string> Coulum_Name, List<string> Coulum_Value)
        {
            bool re = false;
            Exe_Path path = new Exe_Path();

            string acpath = CONNECT_STRING + Exe_Path.exe_Path() + AcPath;
       
           // MessageBox.Show(acpath);


            if (File.Exists(AcPath) == false) { return false; }

            OleDbConnection conect = new OleDbConnection(acpath);

            //---------------------------------------------//

            //string sql = "UPDATE " + TableName + " SET "+String_Joint_Update(Coulum_Name)+" WHERE "+id_Name+"=@a";//条件は[ID]列の値
           // string sql = "UPDATE " + TableName + " SET "+"言語 =@b,名前 =@c,サンプル=@d,ポイント=@e WHERE "+id_Name+" =@a";//条件は[ID]列の値
            string sql = "UPDATE " + TableName +" SET "+String_Joint_Update(Coulum_Name)+ " WHERE " + id_Name + "=@a";//条件は[ID]列の値


            // MessageBox.Show(id_Name+" : "+id_Value);

            conect.Open();

            try
            {
                //MessageBox.Show("Tryの中");

                OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand型変数

                
               // MessageBox.Show("@a  "+ id_Value);


                //command.Parameters.AddWithValue("@b", Coulum_Value[0]);//サンプル
                //command.Parameters.AddWithValue("@c", Coulum_Value[1]);//ポイント
                //command.Parameters.AddWithValue("@a", "2");// Where条件の値【 btn_id 】

                for (int i = 0; i < Coulum_Value.Count; i++)
                {
                   // MessageBox.Show(Query_value[i + 1] + " ,  " + Coulum_Value[i]);
                    command.Parameters.AddWithValue(Query_value[i + 1], Coulum_Value[i]);

                }
                command.Parameters.AddWithValue("@a", id_Value);//Coulum_Value[0]);//条件部分


                int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

               // MessageBox.Show("更新完了：" + count);
                re = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            conect.Close();
            conect.Dispose();
            return re;
        }

        #endregion




        #region テーブルの削除

        public static bool Delete_Table(string AcPath, string TableName) 
        {

            bool re = false;

            string acpath = CONNECT_STRING + AcPath;

            if (File.Exists(AcPath) == false) { return false; }

            OleDbConnection conect = new OleDbConnection(acpath);

            conect.Open();
            try
            {


                string sql = "Drop Table " + TableName;


            OleDbCommand command = new OleDbCommand(sql, conect);  //OleDbCommand　命令を送る



            int count = command.ExecuteNonQuery();//追加件数が返却値（１件登録したら１が返ってくる）

            MessageBox.Show("テーブル削除：" + count);

        
                re = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                re = false;
            }

            conect.Close();
            conect.Dispose();//ガベージコレクション（メモリの開放）

            return re;

        }


        #endregion





        #region 文字列判定

        public static bool is_String(string moji,ListBox list) 
        {

            int num = list.Items.Count;

            while (num > 0)
            {
                list.SelectedIndex = num - 1;

                if (moji == list.Text)
                {
                    return  false;

                }
                num--;
            }
            return true;
        }

        #endregion

    }//クラスの末尾
}
