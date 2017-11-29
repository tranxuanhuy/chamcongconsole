using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab
namespace chamcong
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\myexcel.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string gioquetvantayLythuyet = "";
            bool[] data1rowtungca = new bool[4 * 31];

            //xuat 1 ngay 4 ca la 4 gia tri bool
            for (int i = 10; i <= 10; i++)
            {
                int x=0;
                for (int j = 4; j <= 33; j++)
                {
                    string temp = xlRange.Cells[i, j].Value2.ToString();
                    if (temp.Contains("S")) data1rowtungca[x] = true;
                    if (temp.Contains("T")) data1rowtungca[x+1] = true;
                    if (temp.Contains("C")) data1rowtungca[x+2] = true;
                    if (temp.Contains("D")) data1rowtungca[x+3] = true;
                    x += 4;
                }
                           }
 
            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(@"C:\myexcel.txt", false))
            {
                file.WriteLine(string.Join("\n",data1rowtungca));
            }

            //xuat ra gio dang le phai quet van tay theo ly thuyet
            bool giatridangco = false;
               for (int j = 0; j <124; j++)
               {
                   //truong hop len hoac xuong ca
                   if(data1rowtungca[j]!=giatridangco)
                   {
                       if(data1rowtungca[j]%4==0) 
                   }
               }
        }
    }
}
