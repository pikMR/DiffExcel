using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiffExcel
{
    public static class Configuration
    {
        public static IConfiguration Config { get; set; }
        public static string[] Tables
        {
            get
            {
                return Config.GetSection("Info")["Tables"].Split(";");
            }
        }

        public static string ExcelName
        {
            get
            {
                return Config.GetSection("Excel")["Path"];
            }
        }
    }
}
