using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace insertForm {
    class StringsFile {

       public Dictionary<string, string> dct = new Dictionary<string, string>(); //辞書 【dictionary】 ディクショナリー

       public static string p = System.Windows.Forms.Application.ExecutablePath;

       public static string folderPath = System.IO.Path.GetDirectoryName(p);

	   public static string  FoldaList = folderPath + @"\List\";


        public string AcPth = FoldaList+@"PrevAcPath.txt";

        public string Ac =FoldaList+ @"PrevwDownloadAcName.txt";

        public string Tab=FoldaList+ @"PrevwDownloadTable.txt";

        public string flp = FoldaList+@"PrevwDownloadfilep.txt";

        public string lang= FoldaList+ @"PrevwDownloadLang.txt";

    }
}
