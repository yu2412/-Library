using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace MyCreate
{
   public class File_Move_Copy
    {
        public static string ParentFolder(string FilePath) 
        {
            DirectoryInfo di = new DirectoryInfo(FilePath);
            string ParentPath = di.Parent.FullName;
            return ParentPath;
        }


        public static void Moveing(string Pre, string NextDirect)
        {
       

            if (!File.Exists(Pre))
            {
                MessageBox.Show("ファイルが見つかりません");
                return;
            }


            string FileName = Microsoft.VisualBasic.Interaction.InputBox("変更しない場合は未入力","新しいファイル名");

            if (FileName=="") 
            {

                FileName = System.IO.Path.GetFileName(Pre);

            }

            if (File.Exists(NextDirect+FileName))
            {
                DialogResult dr = MessageBox.Show("同じ名前のデータがあります。\n上書きしますか？\n" + "\r\n" + FileName, "確認", MessageBoxButtons.OKCancel);
                if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    MessageBox.Show("Cancelを押しました。");
                    return;
                }
                else//Cancel以外の動作
                {
                    File.Delete(NextDirect+FileName);

                }

            }

            //PreからNextへフォルダを移動する（両方とも移動させたいファイルの名前が必要）

            // ファイルの移動（同じ名前のファイルがある場合はerror）
            File.Move(Pre, NextDirect+FileName);

            preview(NextDirect);
        }


        public static bool MoveingCpy2(string Prev_PathName, string NextDirect, ref List<string> CpyAcPath)//-------全て設定確認
        {

            if (!File.Exists(Prev_PathName))
            {
                MessageBox.Show("ファイルが見つかりません");
                return false;
            }

                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                //string filePath = Path.GetFileNameWithoutExtension(Prev_PathName) + ".accdb";

                //拡張子ありのファイル名をパスから取得するには
               string FileName = System.IO.Path.GetFileName(Prev_PathName);


            if (File.Exists(NextDirect + FileName))
            {
                  MessageBox.Show("ふぁいるだぶり");
                  return false;
            }

            try
            {
                //PreからNextへフォルダを移動する（両方に移動させたいファイルの名前が必要）
                // ファイルの移動（同じ名前のファイルがある場合はerror）
                File.Copy(Prev_PathName, NextDirect + FileName);
                MessageBox.Show("処理完了");
                //preview(NextDirect);
                CpyAcPath.Add(NextDirect + FileName);
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
        }

        public static bool MoveingCpy1(string Prev_PathName, string NextDirect,ref List<string> CpyAcPath)//-------全て設定確認
        {

            if (!File.Exists(Prev_PathName))
            {
                MessageBox.Show("ファイルが見つかりません");
                return false;
            }


            string FileName = Microsoft.VisualBasic.Interaction.InputBox("変更しない場合は未入力", "新しいファイル名");

            if (FileName == "")
            {
                //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
                //string filePath = Path.GetFileNameWithoutExtension(Prev_PathName) + ".accdb";

                //拡張子ありのファイル名をパスから取得するには
                FileName = System.IO.Path.GetFileName(Prev_PathName);

            }
            else 
            {
                FileName  += ".accdb";
            }

            if (File.Exists(NextDirect + FileName))
            {
                DialogResult dr = MessageBox.Show("同じ名前のデータがあります。\n上書きしますか？\n" + "\r\n" + FileName, "確認", MessageBoxButtons.OKCancel);
                if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    MessageBox.Show("Cancelを押しました。");
                    return false;
                }
                    File.Delete(NextDirect+FileName);//
            }

            try
            {
                //PreからNextへフォルダを移動する（両方に移動させたいファイルの名前が必要）
                // ファイルの移動（同じ名前のファイルがある場合はerror）
                File.Copy(Prev_PathName, NextDirect + FileName);
                MessageBox.Show("処理完了");
                //preview(NextDirect);
                CpyAcPath.Add(NextDirect + FileName);
                return true;
            }
            catch(Exception e) 
            { 
                MessageBox.Show(e.Message);
                return false;
            }
        }

        public static void MoveingCpy_Auto1(string Prev_PathName, string NextDirect)//-------上書き
        {
            if (!File.Exists(Prev_PathName))
            {
                MessageBox.Show("ファイルが見つかりません");
                return;
            }

            string FileName= System.IO.Path.GetFileName(Prev_PathName)+".accdb";


            if (File.Exists(NextDirect + FileName))
            {
                DialogResult dr = MessageBox.Show("同じ名前のデータがあります。\r\n上書きしますか？" + "\r\n" + FileName, "確認", MessageBoxButtons.YesNoCancel);


                if (dr == System.Windows.Forms.DialogResult.Yes)
                {
                    File.Delete(NextDirect + FileName);//
                   // FileName = Microsoft.VisualBasic.Interaction.InputBox("変更しない場合は未入力", "新しいファイル名");
                }
                else 
                {
                    return;
                }


            }

            //PreからNextへフォルダを移動する（両方に移動させたいファイルの名前が必要）
            // ファイルの移動（同じ名前のファイルがある場合はerror）
            File.Copy(Prev_PathName, NextDirect + FileName);

            preview(NextDirect);
        }

        public static void MoveingCpy_Auto2(string Prev_PathName, string NextDirect)//-------スキップ
        {
            if (!File.Exists(Prev_PathName))
            {
                MessageBox.Show("ファイルが見つかりません");
                return;
            }

            string FileName = System.IO.Path.GetFileName(Prev_PathName);


            if (File.Exists(NextDirect + FileName))
            {
                DialogResult dr = MessageBox.Show("同じ名前のデータがあります。\r\n新しいファイル名を設定しますか？" + "\r\n" + FileName, "確認", MessageBoxButtons.YesNoCancel);


                if (dr == System.Windows.Forms.DialogResult.Yes)
                {
                    while (true) {

                        string New_FileName = Microsoft.VisualBasic.Interaction.InputBox("スキップは未入力", "新しいファイル名");
                        if (New_FileName == "") { return; } else if (New_FileName+ ".accdb" == FileName) { MessageBox.Show("名前がかぶっています"); } else { break; }
                    }
                }
                else
                {
                    // File.Delete(NextDirect + FileName);//
                    return;
                }


            }

            //PreからNextへフォルダを移動する（両方に移動させたいファイルの名前が必要）
            // ファイルの移動（同じ名前のファイルがある場合はerror）
            File.Copy(Prev_PathName, NextDirect + FileName);

            preview(NextDirect);
        }



        public static void preview(string Path)
        {


            //エクスプローラでフォルダ"C:\My Documents\My Pictures"を開く
            System.Diagnostics.Process.Start("EXPLORER.EXE", Path);

        }


    }
}
