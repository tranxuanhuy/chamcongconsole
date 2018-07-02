using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;       //Microsoft Excel 14 object in references-> COM tab
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace chamcong
{
    class Program
    {
        private const string fileThucte = "Cham cong Dai Long Security 05.2018.pdf";
        private const string fileLythuyet = "a.xlsx";
        private static int year = 2018;
        private static int month = 5;

        static void Main(string[] args)
        {
            ConvertXLSX.ConvertXLSX2Unicodetxt(@"C:\idnv.xlsx");
            ConvertXLSX.ConvertXLSX2Unicodetxt(fileLythuyet);

            //ConvertXLSX.ConvertXLSX2CSV(@"C:\myexcel1.xlsx");
            File.WriteAllText(System.IO.Path.GetFileNameWithoutExtension(fileThucte),ExtractTextFromPdf(fileThucte));

            List<string> listparam = taoparamconfig();
            foreach (var item in listparam)
            {
                lietkequenchamcong1ng(item);
            }
        }

        public static string ExtractTextFromPdf(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }

                return text.ToString();
            }
        }

        private static List<string> taoparamconfig()
        {
            var data = File.ReadAllLines(@"C:\idnv.txt");
            List<string> listparam = new List<string>();
            foreach (var item in data)
            {
                listparam.Add(item);
            }
            return listparam;
        }

        private static void layhangcuanhanvienlythuyet(List<string> listparam)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileLythuyet);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //lythuyet nv cuoi hang 29, ko co hang trang o sau vn cuoi cung
            int rowCount = 29;
            int colCount = xlRange.Columns.Count;
            for (int k = 0; k < listparam.Count; k++)
            {
                for (int i = 9; i <= rowCount; i++)
                {
                    if (listparam[k].Split(',')[1] == xlRange.Cells[i, 2].Value2.ToString())
                    {
                        listparam[k] += i + ",";
                        break;
                    }
                }
            }

            using (System.IO.StreamWriter file =
          new System.IO.StreamWriter(@"C:\listparam.txt", false))
            {
                file.WriteLine(string.Join("\n", listparam));
            }

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(xlRange);
            Marshal.FinalReleaseComObject(xlWorksheet);

            xlWorkbook.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
        }

        private static void laykhoanghangcuanhanvienthucte(List<string> listparam)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileThucte.Split('.')[0]+".txt");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int k = 0; k < listparam.Count; k++)
            {
                for (int i = 8; i <= rowCount; i++)
                {
                    {
                        if (Convert.ToString(xlRange.Cells[i, 1].Value2) == "1")
                            if (listparam[k].Split(',')[0] == xlRange.Cells[i - 1, 2].Value2.ToString())
                            {
                                listparam[k] += i + ",";
                                for (int j = i; j < rowCount; j++)
                                {
                                    if (Convert.ToString(xlRange.Cells[j, 2].Value2) != null)
                                    {
                                        listparam[k] += j - 1 + ",";
                                        break;
                                    }
                                }
                                break;
                            }
                            else if (listparam[k].Split(',')[0] == Convert.ToString(xlRange.Cells[i - 2, 2].Value2))
                            {
                                listparam[k] += i + ",";
                                for (int j = i; j < rowCount; j++)
                                {
                                    if (Convert.ToString(xlRange.Cells[j, 2].Value2) != null)
                                    {
                                        listparam[k] += j - 1 + ",";
                                        break;
                                    }
                                }
                                break;
                            }
                    }
                }
            }

            using (System.IO.StreamWriter file =
          new System.IO.StreamWriter(@"C:\listparam.txt", false))
            {
                file.WriteLine(string.Join("\n", listparam));
            }

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(xlRange);
            Marshal.FinalReleaseComObject(xlWorksheet);

            xlWorkbook.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
        }

        private static void lietkequenchamcong1ng(string param)
        {
            List<DateTime> gioquetvantayLythuyet = chamconglythuyet(param);
            List<DateTime> gioquetvantayThucte = chamcongthucte(param);
            List<DateTime> cacngayquenchamcong = new List<DateTime>();
            foreach (var lythuyet in gioquetvantayLythuyet)
            {
                bool thuctecochamcong = false;
                foreach (var thucte in gioquetvantayThucte)
                {
                    TimeSpan diff = lythuyet - thucte;
                    double minutes = Math.Abs(diff.TotalMinutes);
                    if (minutes < 70)
                    {
                        thuctecochamcong = true;
                        break;
                    }
                }
                if (!thuctecochamcong) cacngayquenchamcong.Add(lythuyet);
            }
            //xuat ngay gio tho^
            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(@"C:\quenchamcong.txt", true))
            {
                file.WriteLine(param.Split('\t')[1]);
                file.WriteLine(string.Join("\n", cacngayquenchamcong));
            }

            //xuat data report
            // var datalythuyet = File.ReadAllLines(@"C:\chamconglythuyet.txt");
            // using (System.IO.StreamWriter file =
            //new System.IO.StreamWriter(@"C:\dataquenchamcong.txt", true))
            // {
            //     file.WriteLine(param.Split('\t')[1]);
            //     foreach (var ngayquenchamcong in cacngayquenchamcong)
            //     {
            //         foreach (var rowlythuyet in datalythuyet)
            //         {
            //             if (rowlythuyet.Contains(ngayquenchamcong.ToString()))
            //             {
            //                 file.WriteLine(rowlythuyet);
            //                 break;
            //             }
            //         }
            //     }

            // }

            //xuat data report dang doc hieu duoc
            var datalythuyet = File.ReadAllLines(@"C:\chamconglythuyet.txt");
            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(@"C:\dataquenchamcong.txt", true))
            {
                file.WriteLine(param.Split('\t')[1]);
                foreach (var ngayquenchamcong in cacngayquenchamcong)
                {
                    foreach (var rowlythuyet in datalythuyet)
                    {
                        if (rowlythuyet.Contains(ngayquenchamcong.ToString()))
                        {
                            file.WriteLine(reportdungcuphap(rowlythuyet));
                            break;
                        }
                    }
                }

            }
        }

        private static string reportdungcuphap(string rowlythuyet)
        {
            DateTime dt = DateTime.Parse(rowlythuyet.Split(' ')[0]);
            if (rowlythuyet.Split(',')[1] == "x" && rowlythuyet.Split(',')[2] == "d")
            {
                dt = dt.AddDays(-1);
            }
            string data = "Ngày " + dt.ToString("dd/MM/yyyy") + ": ";
            switch (rowlythuyet.Split(',')[1])
            {
                case "l": data += "Lên ca "; break;
                default:
                    data += "Xuống ca ";
                    break;
            }
            switch (rowlythuyet.Split(',')[2])
            {
                case "s": data += "Sáng (00h:06h) "; break;
                case "t": data += "Trưa (06h:12h) "; break;
                case "c": data += "Chiều (12h:18h) "; break;
                default:
                    data += "Đêm (18h:24h) ";
                    break;
            }
            data += "quên chấm công";
            return data;
        }

        private static List<DateTime> chamcongthucte(string param)
        {
            var data = File.ReadAllLines(fileThucte.Split('.')[0] + ".txt");
            List<DateTime> gioquetvantayThucte = new List<DateTime>();

            //lay hang dau tien chua gio cham cong
            bool getrowdata = false;
            string ngaychuadinhdang = "01/11/2017";
            string giochuadinhdang;
            foreach (var item in data)
            {
                if (getrowdata && !item.Contains("No"))
                {
                    //quet hang cuoi cung thi thoat vong lap
                    if (item.Split(',')[0] == "") break;

                    if (!item.Split(',')[0].Contains("\""))
                    {
                        ngaychuadinhdang = item.Split(',')[4];
                        giochuadinhdang = item.Split(',')[5].Replace("\"", "");
                    }
                    else
                        giochuadinhdang = item.Split(',')[0].Replace("\"", "");
                    string[] cacgio = giochuadinhdang.Replace("\n", "").Split(';');
                    foreach (var gio in cacgio)
                    {
                        if (gio != "")
                            gioquetvantayThucte.Add(new DateTime(int.Parse(ngaychuadinhdang.Split('/')[2]), int.Parse(ngaychuadinhdang.Split('/')[1]), int.Parse(ngaychuadinhdang.Split('/')[0]), int.Parse(gio.Split(':')[0]), int.Parse(gio.Split(':')[1]), 0));
                    }
                }
                if (item.Split(',')[1] == param.Split('\t')[0] && !getrowdata) getrowdata = true;
            }

            using (System.IO.StreamWriter file =
          new System.IO.StreamWriter(fileThucte.Split('.')[0] + ".txt", false))
            {
                file.WriteLine(string.Join("\n", gioquetvantayThucte));
            }

            return gioquetvantayThucte;
        }

        private static List<DateTime> chamconglythuyet(string param)
        {
            var data = File.ReadAllLines(fileLythuyet.Split('.')[0] + ".txt");

            bool[] data1rowtungca = new bool[4 * 31 + 1];

            //xuat 1 ngay 4 ca la 4 gia tri bool
            foreach (var row in data)
            {
                if (row.Split('\t')[1] == param.Split('\t')[1])
                {
                    int x = 0;
                    for (int j = 2; j <= 32; j++)
                    {
                        string temp = row.Split('\t')[j];
                        if (temp != null)
                        {
                            if (temp.Contains("S")) data1rowtungca[x] = true;
                            if (temp.Contains("T")) data1rowtungca[x + 1] = true;
                            if (temp.Contains("C")) data1rowtungca[x + 2] = true;
                            if (temp.Contains("D")) data1rowtungca[x + 3] = true;
                        }
                        x += 4;
                    }
                }
            }


            //xuat ra gio dang le phai quet van tay theo ly thuyet
            bool giatridangco = false;
            string[] lenxuongca = new string[4 * 31 + 1];
            List<DateTime> gioquetvantayLythuyet = new List<DateTime>();
            int index = 0;
            //ngay co 30 ngay
            for (int j = 0; j < DateTime.DaysInMonth(year, month) * 4; j++)
            {
                //truong hop len hoac xuong ca
                if (data1rowtungca[j] != giatridangco)
                {
                    giatridangco = !giatridangco;
                    lenxuongca[index] = giatridangco == true ? "l" : "x";
                    //len
                    if (giatridangco)
                    {
                        if (j % 4 == 0) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 0, 00, 0)); lenxuongca[index] += ",s"; }
                        if (j % 4 == 1) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 6, 00, 0)); lenxuongca[index] += ",t"; }
                        if (j % 4 == 2) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 12, 00, 0)); lenxuongca[index] += ",c"; }
                        if (j % 4 == 3) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 18, 00, 0)); lenxuongca[index] += ",d"; }
                    }
                    //xuong                                                  
                    else
                    {
                        if (j % 4 == 0) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 0, 00, 0)); lenxuongca[index] += ",d"; }
                        if (j % 4 == 1) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 6, 00, 0)); lenxuongca[index] += ",s"; }
                        if (j % 4 == 2) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 12, 00, 0)); lenxuongca[index] += ",t"; }
                        if (j % 4 == 3) { gioquetvantayLythuyet.Add(new DateTime(year, month, j / 4 + 1, 18, 00, 0)); lenxuongca[index] += ",c"; }
                    }
                    index++;
                }
            }

            using (System.IO.StreamWriter file =
new System.IO.StreamWriter(@"C:\chamconglythuyet.txt", false))
            {
                int j = 0;
                foreach (var item in gioquetvantayLythuyet)
                {
                    file.WriteLine(item + "," + lenxuongca[j]);
                    j++;
                }

            }

            return gioquetvantayLythuyet;
        }
    }
}