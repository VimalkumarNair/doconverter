using System;
using System.Diagnostics;
using System.IO;

namespace doconverter
{
    class PdfConverter : AbstractConverter, IConverter
    {

        public void toImage(string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            string onlyPath = filePath.Replace(fileName.ToString(), "");

            int resolution = Convert.ToInt32(Configuration.ReadSetting("resolution"));

            string output = onlyPath + this.ExportFolderName;

            if (!Directory.Exists(output))
                Directory.CreateDirectory(output);

            string ghostScriptPath = Configuration.ReadSetting("ghostscript");

            try
            {
                String ars = "-dNOPAUSE -sDEVICE=jpeg -r" + resolution + " -o" + output + "screen%d.jpg -sPAPERSIZE=a4 " + filePath;

                Process proc = new Process();
                proc.StartInfo.FileName = ghostScriptPath;
                proc.StartInfo.Arguments = ars;
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                proc.Start();

                proc.WaitForExit();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
