using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace MyCreate {
    public class Folder {
        static string path = Application.ExecutablePath;
        static string folderPath1 = Path.GetDirectoryName(path);

        #region  ----------実行ファイルパス------------
        public void Explorer_Open() {
            System.Diagnostics.Process.Start("EXPLORER.EXE", folderPath1);
        }

        public void Explorer_Open(string Folder) {
            string folderPath2 = folderPath1 + @"\" + Folder + @"\";
            System.Diagnostics.Process.Start("EXPLORER.EXE", folderPath2);
        }

        public string Exe_Path() {
            string folderPath2 = folderPath1 + @"\";
            return folderPath2;
        }

        public string OpenFile_exe() {


            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "すべてのファイル(*)|*|すべてのファイル(*.*)|*.*";
            op.InitialDirectory = folderPath1 + @"\";

            /* ダイアログを表示し「開く」場合 */
            if (op.ShowDialog() == DialogResult.OK) {


                string Path = op.FileName;//選択したファイルへのパス(文字列型)

                //ファイル名をパスから取得するには、「GetFileNameメソッド」を使います
                string Name = System.IO.Path.GetFileName(Path);
                /* ファイル名を表示する */
                MessageBox.Show(Name);

                return Path;

            } else {
                return "";
            }
        }//------------------

        #endregion

        #region フォルダのパスを獲得
        public string OpenFolder_DaiaLog() //フォルダ指定
        {
            FolderBrowserDialog fbDialog = new FolderBrowserDialog();
            string returnstr = "";

            // ダイアログの説明文を指定する
            fbDialog.Description = "ダイアログの説明文";

            // デフォルトのフォルダを指定する
            fbDialog.SelectedPath = folderPath1;

            // 「新しいフォルダーの作成する」ボタンを表示する
            fbDialog.ShowNewFolderButton = true;

            //フォルダを選択するダイアログを表示する
            if (fbDialog.ShowDialog() == DialogResult.OK) {
                Console.WriteLine(fbDialog.SelectedPath);
                returnstr = fbDialog.SelectedPath;
            } else {
                Console.WriteLine("キャンセルされました");
                returnstr = null;
            }

            // オブジェクトを破棄する
            fbDialog.Dispose();
            return returnstr;
        }

        public string OpenFolder_DaiaLog(string Path) //フォルダ指定
        {
            FolderBrowserDialog fbDialog = new FolderBrowserDialog();
            string returnstr = "";

            // ダイアログの説明文を指定する
            fbDialog.Description = "ダイアログの説明文";

            // デフォルトのフォルダを指定する
            fbDialog.SelectedPath = Path;

            // 「新しいフォルダーの作成する」ボタンを表示する
            fbDialog.ShowNewFolderButton = true;

            //フォルダを選択するダイアログを表示する
            if (fbDialog.ShowDialog() == DialogResult.OK) {
                Console.WriteLine(fbDialog.SelectedPath);
                returnstr = fbDialog.SelectedPath;
            } else {
                Console.WriteLine("キャンセルされました");
                returnstr = null;
            }

            // オブジェクトを破棄する
            fbDialog.Dispose();
            return returnstr;
        }
        #endregion

        public string ParentFolder_Path(string fllpath) {
            return Path.GetDirectoryName(fllpath);//親フォルダまでのパスを取得
        }

        public bool Folder_Exis(string Path) {
            return System.IO.Directory.Exists(Path);
        }

        private void Folder_Exis_Create(string Path) {//フォルダがない場合は作成する
            if (System.IO.Directory.Exists(Path) == false) {

                string CreatePath = Path ;
                System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(CreatePath);
                
            }

        }

        public bool File_Exis(string Path) {
            return System.IO.File.Exists(Path);
        }

        public void Create_Folder_exe(string FolderName) {

            //フォルダ"C:\TEST\SUB"を作成する
            //"C:\TEST"フォルダが存在しなくても"C:\TEST\SUB"が作成される
            //"C:\TEST\SUB"が存在していると、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生

            string CreatePath = folderPath1 + "\\" + FolderName;
            System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(CreatePath);

        }//-------------------------------------------

        public void Create_Folder_Exis(string ParentFolder, string FolderName) {

            //フォルダ"C:\TEST\SUB"を作成する
            //"C:\TEST"フォルダが存在しなくても"C:\TEST\SUB"が作成される
            //"C:\TEST\SUB"が存在していると、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生

            Folder_Exis_Create(ParentFolder);

            string CreatePath = ParentFolder + "\\" + FolderName;
            System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(CreatePath);

        }//-------------------------------------------

        public void Create_Folder(string FolderName) {

            //フォルダ"C:\TEST\SUB"を作成する
            //"C:\TEST"フォルダが存在しなくても"C:\TEST\SUB"が作成される
            //"C:\TEST\SUB"が存在していると、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生

            string DaPath = OpenFolder_DaiaLog();

            string CreatePath = DaPath + "\\" + FolderName;
            System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(CreatePath);


        }//-------------------------------------------

        public void Move_Folder()
        {
            //フォルダ"C:\1"を"C:\2\SUB"に移動（名前を変更）する

            //移動先が別のドライブ（ボリューム）だと、IOExceptionが発生
            //"C:\1"や"C:\2"が存在しないと、DirectoryNotFoundExceptionが発生
            //"C:\1\SUB"のように移動先が移動元のサブフォルダだと、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生



            string beforeFullPath = OpenFolder_DaiaLog();//[～/A]
            string beforePath = Path_Cut.DrectoryName(beforeFullPath);
            MessageBox.Show(beforePath);

            string AfterPath = OpenFolder_DaiaLog() + "\\" + beforePath;//[～/FolderB/A]移動したいフォルダ名に「/」と移動後の「フォルダ名」を足す
            MessageBox.Show(AfterPath);

            try
            {
                Directory.Move(beforeFullPath, AfterPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return;
            }
        }//--------------

        public void Move_Folder(string beforeFullPath)
        {
            //フォルダ"C:\1"を"C:\2\SUB"に移動（名前を変更）する
            //"C:\2\SUB"が存在していると、IOExceptionが発生
            //移動先が別のドライブ（ボリューム）だと、IOExceptionが発生
            //"C:\1"や"C:\2"が存在しないと、DirectoryNotFoundExceptionが発生
            //"C:\1\SUB"のように移動先が移動元のサブフォルダだと、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生

            string beforePath = Path_Cut.DrectoryName(beforeFullPath);
            MessageBox.Show(beforePath);

            string AfterPath = OpenFolder_DaiaLog() + "\\" + beforePath;//[～/FolderB/A]移動したいフォルダ名に「/」と移動後の「フォルダ名」を足す
            MessageBox.Show(AfterPath);

            try
            {
                Directory.Move(beforeFullPath, AfterPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return;
            }
        }//--------------

        public void Move_Folder(string beforePath, string afterPath)
        {
            //フォルダ"C:\1"を"C:\2\SUB"に移動（名前を変更）する
            //"C:\2\SUB"が存在していると、IOExceptionが発生
            //移動先が別のドライブ（ボリューム）だと、IOExceptionが発生
            //"C:\1"や"C:\2"が存在しないと、DirectoryNotFoundExceptionが発生
            //"C:\1\SUB"のように移動先が移動元のサブフォルダだと、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生

            try
            {
                Directory.Move(beforePath, afterPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return;
            }
        }//----------------------------



        public void Delete_Folder(string DeleteFolder) {
            //フォルダ"C:\TEST"を削除する
            //第2項をTrueにすると、"C:\TEST"を根こそぎ（サブフォルダ、ファイルも）削除する
            //"C:\TEST"に読み取り専用ファイルがあると、UnauthorizedAccessExceptionが発生
            //"C:\TEST"が存在しないと、DirectoryNotFoundExceptionが発生
            Directory.Delete(DeleteFolder, true);

        }//------------------------------

    }
}
