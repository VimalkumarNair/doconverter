using System;
using System.IO;
using OfficeCore = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace doconverter
{
    class PowerPointConverter : AbstractConverter, IConverter
    {
        public void toImage(string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            string onlyPath = filePath.Replace(fileName.ToString(), "");

            string output = onlyPath + this.ExportFolderName;

            PowerPoint.Application pptApp;
            PowerPoint.Presentation pptDoc;
            pptApp = new PowerPoint.Application();

            try
            {
                pptDoc = pptApp.Presentations.Open(filePath, OfficeCore.MsoTriState.msoFalse, OfficeCore.MsoTriState.msoFalse, OfficeCore.MsoTriState.msoFalse);

                if (pptDoc.CreateVideoStatus != PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusInProgress)
                    pptDoc.Export(output, "jpg", this.Width, this.Height);

               
                pptDoc.Close();

                //System.Threading.Thread.Sleep(1000);
                pptApp.Quit();
                pptApp = null;

                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
