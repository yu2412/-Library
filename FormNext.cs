using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace insertForm
{
    class FormNextWork
    {
        public enum FormNum : int
        {
            Non=0,
            Form1=1,
            From2=2,
            Form3=3,
            FormSelect=6,
            FormRowMove=9,

        }

        public Form GetForm;

        public string PrevAccessPath;
        public string PrevAccessName;
        public string PrevTableName;

        public string NextAccessPath;
        public string NextAccessName;
        public string NextTableName;

        public Dictionary<string, string> TableRowMenber = new Dictionary<string, string>(); //辞書 【dictionary】 ディクショナリー

       
    }

    class MultiList{
        public List<FormNextWork> nextWorks = new List<FormNextWork>();
    }

}
