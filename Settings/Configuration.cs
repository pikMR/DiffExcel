using Microsoft.Extensions.Configuration;

namespace DiffExcel.Settings
{
    public static class Configuration
    {
        public static IConfiguration Config { get; set; }
        public static string[] Tables => Config.GetSection("Info")["Tables"].Split(";");
        public static string ExcelName => Config.GetSection("Excel")["Path"];
    }
}
