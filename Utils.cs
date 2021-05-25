using System.IO;
using System.Text.RegularExpressions;

namespace DiffExcel
{
    public class Utils
    {
        public static string getRootPath(string rootFilename)
        {
            string _root;
            var rootDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            Regex matchThepath = new(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");
            var appRoot = matchThepath.Match(rootDir).Value;
            _root = Path.Combine(appRoot, rootFilename);
            return _root;
        }
    }
}
