using System.IO;

namespace Doxa.Labs.Excel.Models
{
    public class BaseOutput
    {
        public string Title;
        public string FilePath;
        public string Extension;
        public BaseOutput(string title, string path, Extension extension)
        {
            // set title
            Title = title;

            // set extension
            switch (extension)
            {
                case Models.Extension.Xls:
                    Extension = ".xls";
                    break;
                case Models.Extension.Xlsx:
                    Extension = ".xlsx";
                    break;
            }

            // set path
            FilePath = Path.Combine(path, @"Files\" + title + "" + Extension);
        }
    }
}
