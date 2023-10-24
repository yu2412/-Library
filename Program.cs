using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace insertForm {
    static class Program {


        #region グローバル変数
        public static Form2 F2cntrl;
        public static Form3 F3cntrl;
        public static Form4 F4cntrl;
        public static Form5 F5cntrl;
        public static Form6 F6cntrl;
        public static Form FRow_Movecntrl;
        public static Form FUpDatecntrl;
        public static Form ForcefullyReadCntrl;
        public static txt stdLang;
        public static MultiList MulWork;

        public static StringsFile Prev;
        public static string NewTable_AccessPath ;

        public static string DaleteAccesePath;
  
        public static string SelectAccessname;
        public static string SelectAccessPath;

        public static List<string> SelectTableList;
        public static string SelectTablename;
        public static int SELECT_INDEX = 0;
        public static string[,] AccessRecord;
        public static string[] Save;
        public static RichTextBox txFile;
        public static ComboBox[] combos;
        public static DataGridView daGrVi1;

        #endregion

        public static Form1 MainF;
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MainF = new Form1();
            Application.Run(MainF);
        }
    }
}
