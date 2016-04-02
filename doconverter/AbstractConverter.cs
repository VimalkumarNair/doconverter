using System;

namespace doconverter
{
    abstract class AbstractConverter
    {
        private string exportFolderName;
        private int width;
        private int height;

        /// <summary>
        /// Converter Constructor
        /// </summary>
        public AbstractConverter()
        {
            this.exportFolderName = Configuration.ReadSetting("exportFolderName");

            this.width = Convert.ToInt32(Configuration.ReadSetting("width"));
            this.height = Convert.ToInt32(Configuration.ReadSetting("height"));
        }
                
        public string ExportFolderName
        {
            get { return exportFolderName; }
            set { exportFolderName = value; }
        }

        public int Width
        {
            get { return width; }
            set { width = value; }
        }

        public int Height
        {
            get { return height; }
            set { height = value; }
        }
    }
}
