using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace DiffExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            DocumentTasks dt = new DocumentTasks();
            Config();
            dt.CreateExcel();
        }

        public static void Config()
        {
            Configuration.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.getRootPath("appsettings.json"))
                .Build();

            DbConnectionTarget.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.getRootPath("appsettings.json"))
                .Build();

            DbConnectionSource.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.getRootPath("appsettings.json"))
                .Build();
        }
    }
}
