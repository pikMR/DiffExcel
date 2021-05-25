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
            ConfigurationDB();
            DocumentTasks.CreateExcelWithTableName(DbConnectionTarget.Tables[0]);
        }

        public static void ConfigurationDB()
        {
            DbConnectionTarget.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.getRootPath("appsettings.json"))
                .Build();
        }
    }
}
