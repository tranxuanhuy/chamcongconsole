using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;       //Microsoft Excel 14 object in references-> COM tab
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace chamcong
{
    class Program
    {
        private const string fileThucte = @"C:\Cham cong Dai Long Security 05.2018.pdf";
        private const string fileLythuyet = @"C:\BẢNG CHẤM CÔNG T05 2018 (1).xlsx";
        private static int year = int.Parse(Regex.Match(fileThucte, @"\d{4}").Value);
        private static int month = int.Parse(Regex.Match(fileThucte, @"\d{2}").Value);
        private static string idnv= @"C:\idnv.xlsx";

        static void Main(string[] args)
        {
            File.Delete(@"C:\dataquenchamcong.txt");
            File.Delete(@"C:\quenchamcong.txt");

            ConvertXLSX.ConvertXLSX2Unicodetxt(idnv);
            ConvertXLSX.ConvertXLSX2Unicodetxt(fileLythuyet);

            
            File.WriteAllText(System.IO.Path.GetFileNameWithoutExtension(fileThucte),DateTimeStaffIDFilter(ExtractTextFromPdf(fileThucte)));

            if (IdnvFileHaveAllStaffNameLythuyet() != null)
            {
                Console.WriteLine("ly thuyet");
                Console.WriteLine(IdnvFileHaveAllStaffNameLythuyet());
                Console.ReadKey();
                return;
            }

            if (IdnvFileHaveAllStaffIDThucte()!=null)
            {
                Console.WriteLine("thuc te");
                Console.WriteLine(IdnvFileHaveAllStaffIDThucte());
                Console.ReadKey();
                return; 
            }
            List<string> listparam = taoparamconfig();
            foreach (var item in listparam)
            {
                Console.WriteLine(item.ToString());
                lietkequenchamcong1ng(item);
                
            }
        }

        private static string IdnvFileHaveAllStaffNameLythuyet()
        {
            var stringWithStaffName = File.ReadAllLines("C:\\" + System.IO.Path.GetFileNameWithoutExtension(fileLythuyet) + ".txt").Skip(8);

            string staffNameThieu = null;


            foreach (var item in stringWithStaffName)
            {
                if (!File.ReadAllText("C:\\" + System.IO.Path.GetFileNameWithoutExtension(idnv) + ".txt").Contains(item.Split('\t')[1])&&!string.IsNullOrWhiteSpace(item.Split('\t')[1])&& item.Split('\t')[1].Length<30)
                {
                        staffNameThieu += item.Split('\t')[1] + "\n";
                    
                } 
            }

         
            
            
            return staffNameThieu;
        }

        private static string IdnvFileHaveAllStaffIDThucte()
        {
            string stringWithDateTime = File.ReadAllText(System.IO.Path.GetFileNameWithoutExtension(fileThucte));
        
            string staffIDThieu = null;
            while (true)
            {
     
                Match staffIDMatch = Regex.Match(stringWithDateTime, @"\d{5}");



               if (!string.IsNullOrEmpty(staffIDMatch.Value))
                {
                    string staffIDThucte = staffIDMatch.Value;

                    if (!File.ReadAllText("C:\\"+System.IO.Path.GetFileNameWithoutExtension(idnv)+".txt").Contains(staffIDThucte))
                    {
                        staffIDThieu += staffIDMatch.Value + "\n"; 
                    }


                    stringWithDateTime = stringWithDateTime.Substring(staffIDMatch.Index + 1);


                }
             
                else 
                {
                    break;
                }

            }
            return staffIDThieu;
        }

        private static string DateTimeStaffIDFilter(string stringWithDateTime)
        {
            string filtered = null;
            while (true)
            {
                Match dateMatch = Regex.Match(stringWithDateTime, @"\d{2}\/\d{2}\/\d{4}");
                Match timeMatch = Regex.Match(stringWithDateTime, @"\d{2}\:\d{2}\:\d{2}");
                Match staffIDMatch = Regex.Match(stringWithDateTime, @"\d{5}");

                int minIndex = int.MaxValue;
                if (!string.IsNullOrEmpty(dateMatch.Value))
                    minIndex= Math.Min(minIndex,dateMatch.Index);
                if (!string.IsNullOrEmpty(timeMatch.Value))
                    minIndex = Math.Min(minIndex, timeMatch.Index);
                if (!string.IsNullOrEmpty(staffIDMatch.Value))
                    minIndex = Math.Min(minIndex, staffIDMatch.Index);

                if (minIndex == dateMatch.Index)
                {
                    string date = dateMatch.Value;

                        var dateTime = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.CurrentCulture);

                    //bo qua 23/05/2018
                    //Page 9 of 10No.Employee Full Name Position Record Record Time
                    if (!stringWithDateTime.Substring(dateMatch.Index).Split('\n')[1].StartsWith("Page"))
                    {
                        filtered += dateTime.ToString("dd/MM/yyyy") + "\n"; 
                    }
                        stringWithDateTime = stringWithDateTime.Substring(dateMatch.Index + 1);
                    
                   
                }
               else if (minIndex == timeMatch.Index)
                {
                    string date = timeMatch.Value;

                    var dateTime = DateTime.ParseExact(date, "HH:mm:ss", CultureInfo.CurrentCulture);

                    filtered += dateTime.ToString("HH:mm:ss") + "\n";
                    stringWithDateTime = stringWithDateTime.Substring(timeMatch.Index + 1);


                }
                else if (minIndex == staffIDMatch.Index)
                {
                    string date = staffIDMatch.Value;

                    

                    filtered += date.ToString() + "\n";
                    stringWithDateTime = stringWithDateTime.Substring(staffIDMatch.Index + 1);


                }
                else
                {
                    break;
                }

            }

            
            return filtered;
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

      
      
        private static void lietkequenchamcong1ng(string param)
        {
            List<DateTimeNote> gioquetvantayLythuyet = chamconglythuyet(param);
            List<DateTime> gioquetvantayThucte = chamcongthucte(param);
            List<DateTimeNote> cacngayquenchamcong = new List<DateTimeNote>();
            foreach (var lythuyet in gioquetvantayLythuyet)
            {
                bool thuctecochamcong = false;
                foreach (var thucte in gioquetvantayThucte)
                {
                    TimeSpan diff = lythuyet.DateTime - thucte;
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
                file.WriteLine(string.Join("\n", cacngayquenchamcong.Select(o => o.DateTime).ToList()));
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
                foreach (var item in cacngayquenchamcong)
                {
                    if (!item.Note.Equals("Xuống ca Đêm (18h:24h)"))
                    {
                        file.WriteLine("Ngày " + item.DateTime.ToString("dd/MM/yyyy") + ": " + item.Note + " quên chấm công");
                    }
                    else
                    {
                        file.WriteLine("Ngày " + item.DateTime.AddDays(-1).ToString("dd/MM/yyyy") + ": " + item.Note + " quên chấm công");
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
            string stringWithDateTime = File.ReadAllText(System.IO.Path.GetFileNameWithoutExtension(fileThucte));
            List<DateTime> gioquetvantayThucte = new List<DateTime>();

            try
            {
                stringWithDateTime = stringWithDateTime.Substring(stringWithDateTime.IndexOf(param.Split('\t')[0])+1);
            }
            catch (ArgumentOutOfRangeException)
            {

                return null;
            }
            string dateMatchSave = null ;
            while (true)
            {
                Match dateMatch = Regex.Match(stringWithDateTime, @"\d{2}\/\d{2}\/\d{4}");
                Match timeMatch = Regex.Match(stringWithDateTime, @"\d{2}\:\d{2}\:\d{2}");
                Match staffIDMatch = Regex.Match(stringWithDateTime, @"\d{5}");

                

                int minIndex = int.MaxValue;
                if (!string.IsNullOrEmpty(dateMatch.Value))
                    minIndex = Math.Min(minIndex, dateMatch.Index);
                if (!string.IsNullOrEmpty(timeMatch.Value))
                    minIndex = Math.Min(minIndex, timeMatch.Index);
                if (!string.IsNullOrEmpty(staffIDMatch.Value))
                    minIndex = Math.Min(minIndex, staffIDMatch.Index);

                 if (minIndex == timeMatch.Index)
                {
                    string date = timeMatch.Value;

                    var dateTime = DateTime.ParseExact(date, "HH:mm:ss", CultureInfo.CurrentCulture);

                    gioquetvantayThucte.Add(new DateTime(int.Parse(dateMatchSave.Split('/')[2]), int.Parse(dateMatchSave.Split('/')[1]), int.Parse(dateMatchSave.Split('/')[0]), int.Parse(timeMatch.Value.Split(':')[0]), int.Parse(timeMatch.Value.Split(':')[1]), 0));
                    stringWithDateTime = stringWithDateTime.Substring(timeMatch.Index + 1);


                }
                else if (minIndex == dateMatch.Index)
                {
                    dateMatchSave = dateMatch.Value;
                    stringWithDateTime = stringWithDateTime.Substring(dateMatch.Index + 1);
                }

                else if (minIndex == staffIDMatch.Index|| string.IsNullOrEmpty(staffIDMatch.Value))
                {
                    break;
                }
                
            }

            using (System.IO.StreamWriter file =
          new System.IO.StreamWriter(fileThucte.Split('.')[0] + ".txt", false))
            {
                file.WriteLine(string.Join("\n", gioquetvantayThucte));
            }

            return gioquetvantayThucte;
        }

        private static List<DateTimeNote> chamconglythuyet(string param)
        {
            var data = File.ReadAllLines(fileLythuyet.Split('.')[0] + ".txt");
            List<DateTimeNote> gioquetvantayLythuyet = new List<DateTimeNote>();



            foreach (var row in data)
            {
                if (row.Split('\t')[1] == param.Split('\t')[1])
                {
                    
                    for (int j = 2; j < DateTime.DaysInMonth(year, month)+2; j++)
                    {
                        string temp = row.Split('\t')[j];
                        if (!string.IsNullOrWhiteSpace(temp))
                        {
                            if (temp.Contains("S"))
                            {
                                gioquetvantayLythuyet.Add(new DateTimeNote( new DateTime(year, month, j - 1, 0, 00, 0), "Lên ca Sáng (00h:06h)"));
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 6, 00, 0), "Xuống ca Sáng (00h:06h)"));
                            }
                            if (temp.Contains("T"))
                            {
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 12, 00, 0), "Xuống ca Trưa (06h:12h)"));
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 6, 00, 0), "Lên ca Trưa (06h:12h)"));
                            }
                            if (temp.Contains("C"))
                            {
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 12, 00, 0), "Lên ca Chiều (12h:18h)"));
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 18, 00, 0), "Xuống ca Chiều (12h:18h)"));
                            }
                            if (temp.Contains("D"))
                            {
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 0, 00, 0).AddDays(1), "Xuống ca Đêm (18h:24h)"));
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 18, 00, 0), "Lên ca Đêm (18h:24h)"));
                            }
                            if (temp.Contains("8"))
                            {
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 6, 00, 0).AddDays(1), "Xuống ca 8 (22h:06h)"));
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 22, 00, 0), "Lên ca 8 (22h:06h)"));
                            }
                            if (temp.Contains("3"))
                            {
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 6, 00, 0), "Lên ca 3 (06h:09h)"));
                                gioquetvantayLythuyet.Add(new DateTimeNote(new DateTime(year, month, j - 1, 9, 00, 0), "Xuống ca 3 (06h:09h)"));
                            }

                        }
                        
                    }
                }
            }


            gioquetvantayLythuyet.Sort();
            
                var removeDuplicates = gioquetvantayLythuyet
    .GroupBy(i => i.DateTime)
    .Where(g => g.Count() == 1)
    .Select(g => g.Key);

            List<DateTimeNote> gioquetvantayLythuyetsauxuly = new List<DateTimeNote>();
            foreach (var d in removeDuplicates)
                gioquetvantayLythuyetsauxuly.Add(gioquetvantayLythuyet.Find(item => item.DateTime == d));



            //            using (System.IO.StreamWriter file =
            //new System.IO.StreamWriter(@"C:\chamconglythuyet.txt", false))
            //            {
            //                int j = 0;
            //                foreach (var item in gioquetvantayLythuyet)
            //                {
            //                    file.WriteLine(item + "," + lenxuongca[j]);
            //                    j++;
            //                }

            //            }

            return gioquetvantayLythuyetsauxuly;
        }
    }
}