using System;
using System.IO;
using System.Text.RegularExpressions;

namespace DiffExcel.Settings
{
    public class Utils
    {
        public static string GetRootPath(string rootFilename)
        {
            var rootDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            if (rootDir == null)
                throw new Exception("No se encontró el fichero de configuración appsettings.json");

            Regex matchThepath = new(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");
            var appRoot = matchThepath.Match(rootDir).Value;
            var root = Path.Combine(appRoot, rootFilename);
            return root;
        }
    }
}
