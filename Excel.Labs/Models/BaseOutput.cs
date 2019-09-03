using System.ComponentModel;

namespace Doxa.Labs.Excel.Models
{
    public class BaseOutput
    {
        string Title;
        string Path;
        Extension Extension;
        public BaseOutput(string title, string path, Extension extension)
        {
            Title = title;
            Path = path;
            Extension = extension;
        }
    }

    public enum Extension
    {
        [field: Description(".xls")]
        Xls,
        [field: Description(".xls")]
        Xlsx
    }
}
