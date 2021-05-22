using Microsoft.Extensions.Configuration;
using System;
using System.Data.SqlClient;

namespace DiffExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            ConfigurationDB();
            GetConnectionDBOK();

        }

        public static void ConfigurationDB()
        {
            DbConnection.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.getRootPath("appsettings.json"))
                .Build();
        }

        public static void GetConnectionDBOK()
        {
            
            using (SqlConnection connection = new SqlConnection(
               DbConnection.ConnectionStringTrusted))
            {
                SqlCommand command = new SqlCommand("select * from Person.BusinessEntity", connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            }

        }

    }
}
