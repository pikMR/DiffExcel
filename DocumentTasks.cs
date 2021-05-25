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

            // data of first table
            using SqlConnection connection_first = new(DbConnectionSource.ConnectionStringTrusted);
            using var da_first = new SqlDataAdapter($"select * from {table}", connection_first);
            DataTable datatable_first = new();
            da_first.Fill(datatable_first);

            // data of second table 
            using SqlConnection connection_second = new(DbConnectionTarget.ConnectionStringTrusted);
            using var da_second = new SqlDataAdapter($"select * from {table}", connection_second);
            DataTable datatable_second = new();
            da_second.Fill(datatable_second);
            
            // create Excel
            datatable_first.TableName = $"Source_{table}";
            datatable_second.TableName = $"Target_{table}";
            wb.Worksheets.Add(datatable_first);
            wb.Worksheets.Add(datatable_second);
            wb.SaveAs($"{table}.xlsx");
        }
    }
}
