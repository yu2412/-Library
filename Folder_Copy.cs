using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MyCreate {
   public class Folder_Copy 
   {

        /// <summary>
        /// ディレクトリをコピーする。
        /// </summary>
        /// <param name="sourceDirName">コピーするディレクトリ</param>
        /// <param name="destDirName">コピー先のディレクトリ</param>
        public static void CopyDirectory(string sourceDirName, string destDirName) {
            // コピー先のディレクトリがないかどうか判定する
            if (!Directory.Exists(destDirName)) {
                // コピー先のディレクトリを作成する
                Directory.CreateDirectory(destDirName);
            }

            // コピー元のディレクトリの属性をコピー先のディレクトリに反映する
            File.SetAttributes(destDirName, File.GetAttributes(sourceDirName));

            // ディレクトリパスの末尾が「\」でないかどうかを判定する
            if (!destDirName.EndsWith(Path.DirectorySeparatorChar.ToString())) {
                // コピー先のディレクトリ名の末尾に「\」を付加する
                destDirName = destDirName + Path.DirectorySeparatorChar;
            }

            // コピー元のディレクトリ内のファイルを取得する
            string[] files = Directory.GetFiles(sourceDirName);
            foreach (string file in files) {
                // コピー元のディレクトリにあるファイルをコピー先のディレクトリにコピーする
                File.Copy(file, destDirName + Path.GetFileName(file), true);
            }

            // コピー元のディレクトリのサブディレクトリを取得する
            string[] dirs = Directory.GetDirectories(sourceDirName);
            foreach (string dir in dirs) {
                // コピー元のディレクトリのサブディレクトリで自メソッド（CopyDirectory）を再帰的に呼び出す
                CopyDirectory(dir, destDirName + Path.GetFileName(dir));
            }
        }



    }
}
