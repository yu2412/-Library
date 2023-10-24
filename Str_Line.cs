using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyCreate;

namespace MyCreate
{
  public   class Str_Line
  {
    public static  List<string> String_Add_Line(string moji) 
    {
            List<string> mojiretu = new List<string>();

            //TextBox1に入力されている文字列から一行ずつ読み込む
            //文字列(TextBox1に入力された文字列)からStringReaderインスタンスを作成
            System.IO.StringReader rs = new System.IO.StringReader(moji);

            //StreamReaderを使うと次のようになる
            //System.IO.MemoryStream ms = new System.IO.MemoryStream
            //    (System.Text.Encoding.UTF8.GetBytes(TextBox1.Text));
            //System.IO.StreamReader rs = new System.IO.StreamReader(ms);

            //ストリームの末端まで繰り返す
            while (rs.Peek() > -1)
            {
                //一行読み込んで表示する
                mojiretu.Add(rs.ReadLine());
            }

            return mojiretu;
    }

  }
}
//https://dobon.net/vb/dotnet/string/readline.html#google_vignette