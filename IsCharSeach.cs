using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace insertForm {
    class IsCharSeachList {

        public static bool IsSeach(string BaseStr, ListBox list) {

            int BaseLen = BaseStr.Length;

            int List_Count = list.Items.Count;



            for (int i = 0; i < List_Count; i++) {

                list.SelectedIndex = i;

                string Key = list.Text; ;

                int Key_Len = Key.Length;

                int SeachStart = 0;

                int SeachPoint = 0;


                while (SeachStart + Key_Len - 1 < BaseLen) {

                    SeachPoint = 0;

                    while (BaseStr[SeachStart + SeachPoint] == Key[SeachPoint]) {//検索位置の1文字の照合

                        if (SeachPoint < Key_Len - 1) {
                            SeachPoint += 1;//
                        } else {
                            return true;
                        }

                    }

                    SeachStart += 1;//スタート位置を１つ進めて、検索を０から再スタート



                }
            }
            return false;//文字列中にKeyの文字は入っていなかった
        }

        public static bool IsSeach(string BaseStr, string Key) {

            int BaseLen = BaseStr.Length;

            int Key_Len = Key.Length;

            int SeachStart = 0;

            int SeachPoint = 0;


            while (SeachStart + Key_Len - 1 < BaseLen) {

                SeachPoint = 0;

                while (BaseStr[SeachStart + SeachPoint] == Key[SeachPoint]) {//検索位置の1文字の照合

                    if (SeachPoint < Key_Len - 1) {
                        SeachPoint += 1;//
                    } else {
                        return true;
                    }

                }

                SeachStart += 1;//スタート位置を１つ進めて、検索を０から再スタート



            }

            return false;//文字列中にKeyの文字は入っていなかった

        }





    }
}
