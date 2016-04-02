using System;
using System.Drawing;
using System.IO;

namespace doconverter
{
    class WordConverter : AbstractConverter, IConverter
    {
        public void toImage(string filePath)
        {

            string fileName = Path.GetFileName(filePath);
            string onlyPath = filePath.Replace(fileName.ToString(), "");

            string output = onlyPath + this.ExportFolderName;

            if (!Directory.Exists(output))
                Directory.CreateDirectory(output);

            var docPath = Path.Combine(onlyPath, fileName);
            var app = new Microsoft.Office.Interop.Word.Application();

            //app.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            object nullobj = System.Reflection.Missing.Value;
            object ofalse = false;
            object readOnly = false;

            var doc = new Microsoft.Office.Interop.Word.Document();

            try
            {
                doc = app.Documents.Open(filePath, ref nullobj, ref readOnly,
                                             ref nullobj, ref nullobj, ref nullobj,
                                             ref nullobj, ref nullobj, ref nullobj,
                                             ref nullobj, ref nullobj, ref nullobj,
                                             ref nullobj, ref nullobj, ref nullobj,
                                             ref nullobj);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            doc.Activate();
            doc.Application.Visible = true;


            if (!Directory.Exists(output))
            {
                Directory.CreateDirectory(output);
            }

            try
            {
                //Opens the word document and fetch each page and converts to image
                foreach (Microsoft.Office.Interop.Word.Window window in doc.Windows)
                {
                    foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                    {
                        for (var i = 1; i <= pane.Pages.Count; i++)
                        {
                            var page = pane.Pages[i];

                            var bits = page.EnhMetaFileBits;
                            var target = Path.Combine(output, string.Format("screen" + i.ToString(), i, fileName.Split('.')[0]));

                            using (var ms = new MemoryStream((byte[])(bits)))
                            {
                                var image = System.Drawing.Image.FromStream(ms);

                                var pngTarget = Path.ChangeExtension(target, "jpg");

                                ResizeImage(image, this.Width, this.Height).Save(pngTarget, System.Drawing.Imaging.ImageFormat.Jpeg);

                                System.Threading.Thread.Sleep(500);
                            }
                        }
                    }
                }

                app.Quit(false); // Close Word Application.

                if (doc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                doc = null;
                app = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        /// <summary>
        /// Resizer
        /// </summary>
        /// <param name="image"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            Bitmap result = new Bitmap(width, height);
            result.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            
            using (Graphics graphics = Graphics.FromImage(result))
            {
                graphics.Clear(Color.White);               
                graphics.DrawImage(image, 0, 0, result.Width, result.Height);
            }
            
            return result;
        }
    }
}
