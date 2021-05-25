using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiffExcel
{
    public class DocumentTasks
    {
        public static void CreateExcelWithTableName(string table)
        {
            var wb = new XLWorkbook();
            using SqlConnection connection = new(DbConnectionTarget.ConnectionStringTrusted);
            using var da = new SqlDataAdapter($"select * from {table}", connection);
            DataTable datatable = new();
            da.Fill(datatable);
            datatable.TableName = table;
            wb.Worksheets.Add(datatable);
            wb.SaveAs($"{table}.xlsx");
        }
    }
}
