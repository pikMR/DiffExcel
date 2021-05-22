using Microsoft.Extensions.Configuration;
namespace DiffExcel
{

        public static class DbConnection
        {

            public static IConfiguration Config { get; set; }


            public static string DatabaseName
            {
                get
                {
                    return Config.GetSection("DbConnectionConfig")["DatabaseName"];
                }
            }


            public static string UserName
            {
                get
                {
                    return Config.GetSection("DbConnectionConfig")["UserName"];
                }
            }


            public static string Password
            {
                get
                {
                    return Config.GetSection("DbConnectionConfig")["Password"];
                }
            }


            public static string ServerName
            {
                get
                {
                    return Config.GetSection("DbConnectionConfig")["ServerName"];
                }
            }

            public static string ConnectionString
            {
                get
                {
                    return ($"Server={ServerName};Database={DatabaseName};User ID={UserName};Password={Password};Trusted_Connection=False;MultipleActiveResultSets=true;");
                }
            }

            public static string ConnectionStringTrusted
            {
            //Server=myServerAddress;Database=myDataBase;Trusted_Connection=True;
                get
                {
                    return ($"Server={ServerName};Database={DatabaseName};Trusted_Connection=True;");
                }
            }
    }

}
