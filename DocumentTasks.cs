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
        /// <summary>
        /// Funcionalidad para obtener todas las tablas y para un mismo documento Excel, crear 2 worksheets por cada tabla.
        /// </summary>
        public void CreateExcel()
        {
            var wb = new XLWorkbook();
            //var areEquals = false;
            // obtener todas las tablas
            foreach (var nameTable in Configuration.Tables)
            {
                CreateWorksheets(wb, nameTable);
            }

            wb.SaveAs(Configuration.ExcelName);
        }

        /// <summary>
        /// Create 2 worksheets and add to excel file with data of nameTable
        /// If 2 tables have same Schema then compare and return true if mark diferencess or false if not.
        /// </summary>
        /// <param name="wb">excel file</param>
        /// <param name="nameTable">name of table to compare</param>
        /// <exception cref="System.Exception">Thrown when Schema of tables are diferent</exception>
        private bool CreateWorksheets(XLWorkbook wb, string nameTable)
        {
            // data of first table
            using SqlConnection connection_first = new(DbConnectionSource.ConnectionStringTrusted);
            using var da_first = new SqlDataAdapter($"select * from {nameTable}", connection_first);
            DataTable datatable_first = new();
            da_first.Fill(datatable_first);

            // data of second table 
            using SqlConnection connection_second = new(DbConnectionTarget.ConnectionStringTrusted);
            using var da_second = new SqlDataAdapter($"select * from {nameTable}", connection_second);
            DataTable datatable_second = new();
            da_second.Fill(datatable_second);

            // create 2 new WorkSheets for comparation
            datatable_first.TableName = $"Source_{nameTable}";
            datatable_second.TableName = $"Target_{nameTable}";
            wb.Worksheets.Add(datatable_first);
            wb.Worksheets.Add(datatable_second);

            // Compare last
            var compareWS = wb.Worksheets.TakeLast(2);
            var ws1 = compareWS.ElementAt(0);
            var ws2 = compareWS.ElementAt(1);

            if (ws1.ColumnCount() != ws2.ColumnCount())
                throw new Exception($"El número de columnas es diferente para el worksheet target {ws1.Name} y el worksheet source {ws2.Name} de la tabla {nameTable}");

            return PaintDifferences(ws1,ws2);
        }

        private bool PaintDifferences(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            throw new NotImplementedException();
        }
    }
}
