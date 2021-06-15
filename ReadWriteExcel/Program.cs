using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace ReadWriteExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            string filePath = @"";
            //ReadExcelSample(filePath);
            //WriteExcelSample(filePath);

            Console.Read();
        }

        public static void ReadExcelSample(string filePath) 
        {
            Application app = new Application();
            try 
            {
                Workbook workbook = app.Workbooks.Open(filePath);
                Worksheet worksheet = workbook.Sheets[1];
                Range range = worksheet.UsedRange;
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;

                for (int i = 1; i < rows; i++)
                {
                    for (int j = 1; j < cols; j++)
                    {
                        Console.WriteLine(range.Cells[i, j].Value2.ToString());
                    }
                }
            }
            catch(Exception ex) 
            {
                Console.WriteLine($"讀取Excel發生例外:{ex.Message}\r\n{ex.StackTrace}");
            }
            finally 
            {
                app.Quit();
            }
        }

        public static void WriteExcelSample(string filePath) 
        {
            Application app = new Application();
            try
            {
                Workbook workbook = app.Workbooks.Open(filePath);
                Worksheet worksheet = workbook.Sheets[1];
                Range range = worksheet.UsedRange;
                int rows = range.Rows.Count;
                int cols = range.Columns.Count;


                range.Cells[1, 1].Value2 = "123456";
                workbook.Save();


            }
            catch (Exception ex)
            {
                Console.WriteLine($"寫入Excel發生例外:{ex.Message}\r\n{ex.StackTrace}");
            }
            finally
            {
                app.Quit();
            }
        }
    }
}
