using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MyCreate{
   public class Path_Cut {

        public static string FileData(string path) {///拡張子ありのファイル名

            //ファイル名をパスから取得するには、「GetFileNameメソッド」を使います
            string filePath = Path.GetFileName(path);
            return filePath;
        }

        public static string File_Extension(string path) {///拡張子だけ
            //拡張子をパスから取得するには、「GetExtensionメソッド」を使います。
            string filePath = Path.GetExtension(path);
            return filePath;
        }

        public static string FileName(string path) {///拡張子なしのファイル名

            //拡張子なしのファイル名をパスから取得するには、「GetFileNameWithoutExtensionメソッド」を使います。
            string filePath = Path.GetFileNameWithoutExtension(path);
            return filePath;
        }

        public static string DrectoryName(string path) { ///指定ファイルパスの親フォルダの名前を取得
            //ディレクトリ名をパスから取得するには、「GetDirectoryNameメソッド」を使います。
            string name = Path.GetDirectoryName(path);
            return name;
        }

    }
}
