using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace insertForm
{
    class Drag_Drop
    {

        public static void listCount(List<string> Drlist, ListBox langlist, ListBox list_0, ListBox list_00, ListBox NGlist)
        {
            list_0.Items.Clear();
            list_00.Items.Clear();


            if (langlist.Text == "")
            {
                MessageBox.Show("言語を選択してください");

            }

            int num = Drlist.Count;

            MessageBox.Show(num + "個");

            while (num > 0)
            {


                if (Directory.Exists(Drlist[num - 1]))
                {

                    MessageBox.Show("ふぉるだです");
                }
                else if (File.Exists(Drlist[num - 1]))
                {
                    MessageBox.Show("ファイルです");

                    //ファイル名をパスから取得するには、「GetFileNameメソッド」を使います
                    string Search = Path.GetFileName(Drlist[num-1]);
                    list_0.Items.Add(Drlist[num - 1]);
                    list_00.Items.Add(Search);

                    Drlist.RemoveAt(num - 1);

                    num = Drlist.Count;
                    continue;

                }


                string strp = Drlist[num - 1];
                string Fipth = Drlist[num - 1];

                //----------------------------------------------//


                string Gengo = langlist.Text;

                //IEnumerable<string> files;

                Object a;

                switch (Gengo)
                {


                    case "C#":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.cs", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly


                        break;



                    case "Java":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.java", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly


                        break;



                    case "PHP":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.php", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                        break;



                    case "SQL":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.txt", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                        break;


                    case "Excel_VBA":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.txt", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                        break;



                    case "Access_VBA":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.txt", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                        break;



                    case "C言語":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.c", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                        break;

                    case "C++":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.cpp", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                        break;

                    case "JavaScript":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.js", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly
                        break;

                    case "HTML":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.html", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly
                        break;

                    case "CSS":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.css", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly
                        break;

                    case "テキスト":
                        a = System.IO.Directory.EnumerateFiles(strp, "*.txt", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly
                        break;

                    default:
                        a = System.IO.Directory.EnumerateFiles(strp, "*", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly
                        break;

                }
                IEnumerable<string> files = (IEnumerable<string>)a;

                MessageBox.Show("File数："+files.Count());


                //ファイルを列挙する
                foreach (string f in files)
                {




                    ////拡張子なしのファイル名をパスから取得する
                    //string txtname = Path.GetFileNameWithoutExtension(f);


                    //ファイル名をパスから取得するには、「GetFileNameメソッド」を使います
                    string Search = Path.GetFileName(f);




                    switch (Gengo)
                    {


                        case "C#":
                            if (IsCharSeachList.IsSeach(Search, NGlist) == false)
                            {
                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            }


                            break;



                        case "Java":
                            if (IsCharSeachList.IsSeach(Search, NGlist) == false)
                            {
                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            }


                            break;



                        case "PHP":
                            if (IsCharSeachList.IsSeach(Search, NGlist) == false)
                            {
                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            }

                            break;



                        case "SQL":
                            if (IsCharSeachList.IsSeach(Search, NGlist) == false)
                            {
                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            }
                            break;


                        case "C言語":

                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            
                            break;

                        case "C++":


                            list_0.Items.Add(f);
                            list_00.Items.Add(Search);


                            break;



                        case "text":

  
                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            

                            break;


                        default:
                                list_0.Items.Add(f);
                                list_00.Items.Add(Search);
                            

                            break;

                    }

                }//-----foreach

                 Drlist.RemoveAt(num - 1);
                MessageBox.Show("処理終了後：" + Drlist.Count + " 個");
                num = Drlist.Count;


            }
        }

        //-----------------------------------------------------------------------------//


        public static void listCount(List<string> Drlist, ListBox list_0, ListBox list_00) {//Access用のやつ
            list_0.Items.Clear();
            list_00.Items.Clear();




            int num = Drlist.Count;

            MessageBox.Show(num + "個");

            while (num > 0) {


                if (Directory.Exists(Drlist[num - 1])) {

                    MessageBox.Show("ふぉるだです");



                } else if (File.Exists(Drlist[num - 1])) {
                    MessageBox.Show("ファイルです");

                    //ファイル名をパスから取得するには、「GetFileNameメソッド」を使います
                    string Search = Path.GetFileName(Drlist[num - 1]);
                    list_0.Items.Add(Drlist[num - 1]);
                    list_00.Items.Add(Search);

                    Drlist.RemoveAt(num - 1);

                    num = Drlist.Count;
                    continue;

                }


                string strp = Drlist[num - 1];
                string Fipth = Drlist[num - 1];

                //----------------------------------------------//

                Object a;
                 a = System.IO.Directory.EnumerateFiles(strp, "*.accdb", System.IO.SearchOption.AllDirectories);//TopDirectoryOnly

                IEnumerable<string> files = (IEnumerable<string>)a;

                MessageBox.Show("File数：" + files.Count());


                //ファイルを列挙する
                foreach (string f in files) {


                    //ファイル名をパスから取得するには、「GetFileNameメソッド」を使います
                    string Search = Path.GetFileName(f);
                    

                    list_0.Items.Add(f);
                            list_00.Items.Add(Search);

                }//-----foreach

                Drlist.RemoveAt(num - 1);
                MessageBox.Show("処理終了後：" + Drlist.Count + " 個");
                num = Drlist.Count;


            }
        }


    }
}