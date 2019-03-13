using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace textgen
{
    static class DocProcessingcs
    {
        public static void DocHandler()
        {
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\User\source\repos\textgen\textgen\Properties\dataset-orig.xlsx", 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); 
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Range range = xlWorkSheet.UsedRange;

            string str = ""; // variable for current text
            int rw = range.Rows.Count;
            int cl = 7; //the main colomn for our task
            int num = 1; //the indexer for gerenerated files

            for (int row = 2; row < rw; row++)
            {
                str = (string)(range.Cells[row, cl] as Range).Value2;
                CreateTXT(str, num);
                num++;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public static void CreateTXT(string value, int num)
        {
            string path = (@"C:\data-output\") + num + ".txt";
            StreamWriter textFile = new StreamWriter(path);
            textFile.Write(value);
            textFile.Close();
        }
    }
}
