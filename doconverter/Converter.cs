namespace doconverter
{
    enum FileType
    {
        PDF,
        Excel,
        Word,
        Powerpoint
    };

    class Converter
    {
        FileType type;

        public Converter(FileType type)
        {
            this.Type = type;
        }

        public void ConvertToImage(string filePath)
        {
            if (this.type == FileType.Excel)
            {
                new ExcelConverter().toImage(filePath);
            }
            else if (this.type == FileType.Word)
            {
                new WordConverter().toImage(filePath);
            }
            else if (this.type == FileType.Powerpoint)
            {
                new PowerPointConverter().toImage(filePath);
            }
            else if (this.type == FileType.PDF)
            {
                new PdfConverter().toImage(filePath);
            }
        }

        protected FileType Type
        {
            get { return type; }
            set { type = value; }
        }
    }
}
