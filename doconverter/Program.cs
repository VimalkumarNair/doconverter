using System;
using System.IO;
using System.Windows.Forms;
using System.Configuration;

namespace doconverter
{
    class Program
    {
        static string filePath;

        [STAThread]
        static void Main(string[] args)
        {
            filePath = args[0];
            
            string fileExt = Path.GetExtension(filePath).Trim('.');

            try
            {
                switch (fileExt)
                {
                    case "xls":
                    case "xlsx":
                        new Converter(FileType.Excel).ConvertToImage(filePath);
                        break;
                    case "doc":
                    case "docx":
                        new Converter(FileType.Word).ConvertToImage(filePath);
                        break;
                    case "ppt":
                    case "pptx":
                        new Converter(FileType.Powerpoint).ConvertToImage(filePath);
                        break;
                    case "pdf":
                        new Converter(FileType.PDF).ConvertToImage(filePath);
                        break;
                    default:
                        Console.WriteLine("Unsupported file format!");
                        break;
                }

                Console.Write("Success");

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }

            Environment.Exit(0);
        }
    }
}
