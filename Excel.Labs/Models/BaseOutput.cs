namespace Doxa.Labs.Excel.Models
{
    public class BaseOutput
    {
        public string Title;
        public string Path;
        public string Extension;
        public BaseOutput(string title, string path, Extension extension)
        {
            Title = title;
            Path = path;

            // set extension
            switch (extension)
            {
                case Models.Extension.Xls: Extension = ".xls";
                    break;
                case Models.Extension.Xlsx: Extension = ".xlsx";
                    break;
            }
        }
    }
}
