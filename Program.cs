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
            GetDataDBOK();
        }

        public static void ConfigurationDB()
        {
            DbConnectionTarget.Config = new ConfigurationBuilder()
                .AddJsonFile(Utils.getRootPath("appsettings.json"))
                .Build();
        }

        public static void GetConnectionDBOK()
        {
            using (SqlConnection connection = new SqlConnection(
               DbConnectionTarget.ConnectionStringTrusted))
            {
                SqlCommand command = new SqlCommand("select * from Person.BusinessEntity", connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
        }

        public static void GetDataDBOK()
        {
            SqlDataReader rdr = null;

            try
            {
                using (SqlConnection connection = new SqlConnection(
               DbConnectionTarget.ConnectionStringTrusted))
                {
                    SqlCommand command = new SqlCommand("select * from Person.BusinessEntity", connection);
                    command.Connection.Open();
                    rdr = command.ExecuteReader();
                    // print the CustomerID of each record
                    while (rdr.Read())
                    {
                        Console.WriteLine(rdr[0]);
                    }
                    //command.ExecuteNonQuery();
                }
            }
            finally
            {
                // close the reader
                if (rdr != null)
                {
                    rdr.Close();
                }
            }
        }

    }
}
