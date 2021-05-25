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
            Config();
            DocumentTasks.CreateExcelWithTableName(Configuration.Tables[0]);
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
