using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
        //
// OfficeアプリのEXEパスを取得するライブラリ
//
// This software is distributed under the license of NYSL.
// http://www.kmonos.net/nysl/
// 
//
namespace MyCreate
{
    /// <summary>
    /// Officeアプリの種類のenum型
    /// </summary>
    public enum Office_Num
    {
        /// <summary>Word</summary>
        WORD,
        /// <summary>Excel</summary>
        EXCEL,
        /// <summary>PowerPoint</summary>
        POWERPOINT,
        /// <summary>Outlook</summary>
        OUTLOOK,
        /// <summary>OneNote</summary>
        ONENOTE,
        /// <summary>Access</summary>
        ACCESS,
        /// <summary>Publisher</summary>
        PUBLISHER
    }



    /// <summary>
    /// OfficeアプリのEXEパスを取得するライブラリ
    /// </summary>
    public class MicrosoftApp
   {
            /// <summary>
            /// OfficeアプリのEXEファイル名を取得する
            /// </summary>
            /// <param name="app">Officeアプリの種類</param>
            /// <returns>EXEファイル名</returns>
            /// <exception cref="Exception"></exception>
            public static String Get_OfficeAppExeName(Office_Num app)
            {
                switch (app)
                {
                    case Office_Num.WORD:
                        return "WINWORD.EXE";
                    case Office_Num.EXCEL:
                        return "excel.exe";
                    case Office_Num.POWERPOINT:
                        return "POWERPNT.EXE";
                    case Office_Num.OUTLOOK:
                        return "OUTLOOK.EXE";
                    case Office_Num.ONENOTE:
                        return "ONENOTE.EXE";
                    case Office_Num.ACCESS:
                        return "MSACCESS.EXE";
                    case Office_Num.PUBLISHER:
                        return "mspub.exe";
                }

                throw new Exception();
            }

            /// <summary>
            /// OfficeアプリのEXEファイルのパスを取得する
            /// </summary>
            /// <param name="app">Officeアプリの種類</param>
            /// <returns>OfficeアプリのEXEファイルパス</returns>
            /// <exception cref="Exception"></exception>
            public static String Get_OfficeAppPath(Office_Num app)
            {
                // EXEファイル名を決定
                var exeName = Get_OfficeAppExeName(app);
                // EXEファイルに対応するレジストリキーを決定
                var subkeyName = $@"Software\Microsoft\Windows\CurrentVersion\App Paths\{exeName}";

                // レジストリからEXEファイルのパスを取得し返す
                using (var subkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(subkeyName, false))
                {
                    if (subkey == null)
                    {
                        throw new Exception();
                    }

                    return subkey.GetValue(string.Empty)?.ToString();
                }
            }
    }
 }


