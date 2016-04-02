using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows.Forms;

namespace doconverter
{
    class ExcelConverter : AbstractConverter, IConverter
    {

        public void toImage(string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            string onlyPath = filePath.Replace(fileName.ToString(), "");

            string output = onlyPath + this.ExportFolderName;

            if (!Directory.Exists(output))
                Directory.CreateDirectory(output);

            Excel.Application oExcel = new Excel.Application();
            Excel.Workbook wb = null;
            oExcel.DisplayAlerts = false;

            try
            {
                wb = oExcel.Workbooks.Open(filePath.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                Excel.Sheets sheets = wb.Worksheets as Excel.Sheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel.Worksheet sheet = sheets[i];
                    string startRange = "A1";
                    //Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            
                    string endRange = "AP50";
                    Excel.Range range = sheet.get_Range(startRange, endRange);
                    range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                    Image imgRange = this.GetImageFromClipboard();

                    imgRange.Save(output + i.ToString() + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                    wb.Save();

                }

                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(new Proccessor().GetExcelProcess(oExcel).Id);

                p.Kill();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public Image GetImageFromClipboard()
        {
            Image returnImage = null;

            IDataObject d = Clipboard.GetDataObject();
            if (d.GetDataPresent(DataFormats.Bitmap))
            {
                returnImage = Clipboard.GetImage();
            }

            return returnImage;
        }
    }
}
