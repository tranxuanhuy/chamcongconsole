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
            //chamconglythuyet();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\myexcel1.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<DateTime> gioquetvantayThucte = new List<DateTime>();
            for (int i = 9; i <= 28; i++)
            {
                string ngaychuadinhdang = xlRange.Cells[i, 5].Value2.ToString();
                double date = double.Parse(ngaychuadinhdang);

                var ngay = DateTime.FromOADate(date).ToString("dd/MM/yyyy");
                string gio = xlRange.Cells[i, 6].Value2.ToString();
                string[] cacgio = gio.Replace("\n","").Split(';');
                foreach (var item in cacgio)
                {
                    gioquetvantayThucte.Add(new DateTime(int.Parse(ngay.Split('/')[2]), int.Parse(ngay.Split('/')[1]), int.Parse(ngay.Split('/')[0]), int.Parse(item.Split(':')[0]), int.Parse(item.Split(':')[1]), 0)); 
                }
             }
            using (System.IO.StreamWriter file =
          new System.IO.StreamWriter(@"C:\myexcel1.txt", false))
            {
                file.WriteLine(string.Join("\n", gioquetvantayThucte));
            }
        }

        private static void chamconglythuyet()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\myexcel.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            bool[] data1rowtungca = new bool[4 * 31];

            //xuat 1 ngay 4 ca la 4 gia tri bool
            for (int i = 10; i <= 10; i++)
            {
                int x = 0;
                for (int j = 4; j <= 33; j++)
                {
                    string temp = xlRange.Cells[i, j].Value2.ToString();
                    if (temp.Contains("S")) data1rowtungca[x] = true;
                    if (temp.Contains("T")) data1rowtungca[x + 1] = true;
                    if (temp.Contains("C")) data1rowtungca[x + 2] = true;
                    if (temp.Contains("D")) data1rowtungca[x + 3] = true;
                    x += 4;
                }
            }

            //xuat ra gio dang le phai quet van tay theo ly thuyet
            bool giatridangco = false;
            List<DateTime> gioquetvantayLythuyet = new List<DateTime>();
            for (int j = 0; j < 124; j++)
            {
                //truong hop len hoac xuong ca
                if (data1rowtungca[j] != giatridangco)
                {
                    giatridangco = !giatridangco;
                    if (j % 4 == 0) gioquetvantayLythuyet.Add(new DateTime(2017, 11, j / 4 + 1, 0, 00, 0));
                    if (j % 4 == 1) gioquetvantayLythuyet.Add(new DateTime(2017, 11, j / 4 + 1, 6, 00, 0));
                    if (j % 4 == 2) gioquetvantayLythuyet.Add(new DateTime(2017, 11, j / 4 + 1, 12, 00, 0));
                    if (j % 4 == 3) gioquetvantayLythuyet.Add(new DateTime(2017, 11, j / 4 + 1, 18, 00, 0));
                }
            }

            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(@"C:\myexcel.txt", false))
            {
                file.WriteLine(string.Join("\n", gioquetvantayLythuyet));
            }
        }
    }
}
