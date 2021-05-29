using Microsoft.Extensions.Configuration;
namespace DiffExcel.Settings
{
    public static class DbConnectionSource
        {
            private const string JsonDbConnectionConfig = "DbConnectionConfigSource";
            public static IConfiguration Config { get; set; }
            public static string DatabaseName => Config.GetSection(JsonDbConnectionConfig)["DatabaseName"];
            public static string UserName => Config.GetSection(JsonDbConnectionConfig)["UserName"];
            public static string Password => Config.GetSection(JsonDbConnectionConfig)["Password"];
            public static string ServerName => Config.GetSection(JsonDbConnectionConfig)["ServerName"];

            public static string ConnectionString => ($"Server={ServerName};Database={DatabaseName};User ID={UserName};Password={Password};Trusted_Connection=False;MultipleActiveResultSets=true;");

            public static string ConnectionStringTrusted => ($"Server={ServerName};Database={DatabaseName};Trusted_Connection=True;");

            public static bool IsTrusted => bool.Parse(Config.GetSection(JsonDbConnectionConfig)["Trusted"]);
        }

}
