using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
namespace chamcong
{
    class ConvertXLSX
    {
        public static void ConvertXLSX2Unicodetxt(string args)
        {
            Excel.Application app = new Excel.Application();

            try
            {
                app.DisplayAlerts = false;
                app.Visible = false;

                Excel.Workbook book = app.Workbooks.Open(args);

                book.SaveAs(Filename: args.Replace("xlsx","txt"), FileFormat: Excel.XlFileFormat.xlUnicodeText,
                    AccessMode: Excel.XlSaveAsAccessMode.xlNoChange,
                    ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            }
            finally
            {
                app.Quit();
            }
        }

        public static void ConvertXLSX2CSV(string args)
        {
            Excel.Application app = new Excel.Application();

            try
            {
                app.DisplayAlerts = false;
                app.Visible = false;

                Excel.Workbook book = app.Workbooks.Open(args);

                book.SaveAs(Filename: args.Replace("xlsx", "csv"), FileFormat: Excel.XlFileFormat.xlCSV,
                    AccessMode: Excel.XlSaveAsAccessMode.xlNoChange,
                    ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            }
            finally
            {
                app.Quit();
            }
        }
    }
}
