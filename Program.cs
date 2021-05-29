using DiffExcel.Excel;
using DiffExcel.Settings;
using Microsoft.Extensions.Configuration;

namespace DiffExcel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var dt = new DocumentTasks();
            Config();
            dt.CreateExcel();
        }

        public static void Config()
        {
            Configuration.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.GetRootPath("appsettings.json"))
                .Build();

            DbConnectionTarget.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.GetRootPath("appsettings.json"))
                .Build();

            DbConnectionSource.Config = new ConfigurationBuilder()
                // ReSharper disable once StringLiteralTypo
                .AddJsonFile(Utils.GetRootPath("appsettings.json"))
                .Build();
        }
    }
}
