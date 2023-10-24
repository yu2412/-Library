using insertForm;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MyCreate
{
    class ExcelManager : IDisposable
    {


        /// <summary>
        /// Excel操作用オブジェクト
        /// </summary>
        private Microsoft.Office.Interop.Excel.Application _application = null;
        private Workbook _workbook = null;
        private Worksheet _worksheet = null;
        public Microsoft.Office.Interop.Excel.Sheets sheets = null;


        //static string[] strLine = new string[500];
        // static string[] TestLine = new string[] { "aaa", "bbb", "ccc", "ddd" };//テスト用

        // Dispose処理は、下記のページを参考にしています。
        // [アンマネージドリソースをDisposeパターンで管理する]
        // (https://days-of-programming.blogspot.com/2018/04/dispose.html)


        public static void OPEN_Excell(string ex_Path)
        {
            //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE");
           
            Office_Num office = Office_Num.EXCEL;
           string ex= MicrosoftApp.Get_OfficeAppPath(office);

            System.Diagnostics.Process.Start(ex);

            var proc = new System.Diagnostics.Process();

            proc.StartInfo.FileName = ex_Path;
            proc.Start();


        }






        #region "IDisposable Support"
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: Managed Objectの破棄
                }

                if (_workbook != null)
                {
                    _workbook.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }

                if (_application != null)
                {
                    _application.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);
                    _application = null;
                }

                disposedValue = true;
            }
        }

        ~ExcelManager()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion "IDisposable Support"

        /// <summary>
        /// Excelワークブックを開く
        /// </summary>
        public void Open(string strPath,string SheetName)
        {
            try
            {
                // Excelアプリケーション生成
                _application = new Microsoft.Office.Interop.Excel.Application()
                {
                    // 非表示
                    //Visible = true
                    Visible = false
                };

                // Bookを開く
                _workbook = _application.Workbooks.Open(strPath);

                // 対象シートを設定する
                _worksheet = _workbook.Worksheets[SheetName];//
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        /// <summary>
        /// Excelワークブックをファイル名を指定して保存する
        /// </summary>
        /// <returns>True:正常終了、False:保存失敗</returns>
        public bool SaveAs()
        {
            try
            {
                // ファイル名を指定して保存する
                _workbook.SaveCopyAs(@"e:\book2.xlsx");
            }
            catch
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// セル設定
        /// </summary>
        /// <param name="rowIndex">row</param>
        /// <param name="columnIndex">col</param>
        /// <param name="value">value</param>
        public void WriteCell(int rowIndex, int columnIndex, object value)
        {
            // セルを指定
            var cells = _worksheet.Cells;
            var range = cells[rowIndex, columnIndex] as Range;

            // 値を設定
            range.Value = value;

            // cell解放
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(cells);
        }

        /*  --------------------------------------------------------------------------------------------  */

        public static void Excel_Write(string Path ,List<string> List)
        {

            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbooks books = null;
            Microsoft.Office.Interop.Excel.Workbook book = null;
            Microsoft.Office.Interop.Excel.Sheets sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            Microsoft.Office.Interop.Excel.Range cells = null;
            Microsoft.Office.Interop.Excel.Range range = null;

        try
        {
                //Excelファイルの場所
                string file_path = Path;


                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.DisplayAlerts = false; // アラートを表示しない
                books = excel.Workbooks;
                book = books.Open(file_path); // Excelを開く (excel.Workbooks.Open() とすると解放漏れが起こるので×です)


                ///*********///
                //Excelファイル操作処理～～
                sheets = book.Worksheets;
                int sheetnumber = sheets.Count;

                var sheetEnd = excel.Worksheets[sheetnumber];

              //  sheets.Add(After: sheetEnd);
               // sheetnumber = sheets.Count;//再カウント

                //excel.Worksheets[sheetnumber].name = "新しいシートおお";
              //  excel.Worksheets[sheetnumber].name = "テストsheet";//作成シート名



                sheet = sheets[sheetnumber]; // 最初のシート。book.Worksheets[1] とするのも×です。
                cells = sheet.Cells; // セル一覧のオブジェクト。

                for (int i = 0; i < List.Count; i++)
                {
                    sheet.Cells[i + 4, 3] = List[i]; // 左上のセル(Range)。sheet.Cells[1, 1] とするのも×です。
                }

                book.Close(SaveChanges: true); // 保存して閉じる
                excel.Quit(); // Excelを終了する

                MessageBox.Show("処理が完了しました。");
            } ///*********///
            catch (Exception)
            {
                if (book != null) book.Close(SaveChanges: false); // 閉じていなければ閉じる
                if (excel != null) excel.Quit(); // 終了していなければ終了

                MessageBox.Show("処理が失敗しました。");
            }
            finally
            {

                // 解放漏れがないようにしてください。
                FinalReleaseComObject(range);
                FinalReleaseComObject(cells);
                FinalReleaseComObject(sheet);
                FinalReleaseComObject(sheets);
                FinalReleaseComObject(book);
                FinalReleaseComObject(books);
                FinalReleaseComObject(excel);
            }
        }

        private static void FinalReleaseComObject(object o)
        {
            if (o != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
        }


    }
}
