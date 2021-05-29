using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace DiffExcel
{
    public class DocumentTasks
    {
        private string TableName {get;set;}

        /// <summary>
        /// Funcionalidad para obtener todas las tablas y para un mismo documento Excel, crear 2 worksheets por cada tabla.
        /// </summary>
        public void CreateExcel()
        {
            var wb = new XLWorkbook();

            foreach (var nameTable in Configuration.Tables)
            {
                TableName = nameTable;
                var listDiff = CreateWorksheets(wb, nameTable);
                listDiff.ForEach(x => Console.WriteLine(x.ToString()));
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
        private List<DiferenceWs> CreateWorksheets(XLWorkbook wb, string nameTable)
        {
            // data of first table
            using SqlConnection connectionFirst = new(DbConnectionSource.ConnectionStringTrusted);
            using var daFirst = new SqlDataAdapter($"select * from {nameTable}", connectionFirst);
            DataTable datatableFirst = new();
            daFirst.Fill(datatableFirst);

            // data of second table 
            using SqlConnection connectionSecond = new(DbConnectionTarget.ConnectionStringTrusted);
            using var daSecond = new SqlDataAdapter($"select * from {nameTable}", connectionSecond);
            DataTable datatableSecond = new();
            daSecond.Fill(datatableSecond);

            // create 2 new WorkSheets for comparation
            datatableFirst.TableName = $"Source_{nameTable}";
            datatableSecond.TableName = $"Target_{nameTable}";
            wb.Worksheets.Add(datatableFirst);
            wb.Worksheets.Add(datatableSecond);

            // Compare last
            var compareWs = wb.Worksheets.TakeLast(2);
            var ws1 = compareWs.ElementAt(0);
            var ws2 = compareWs.ElementAt(1);

            return PaintDifferences(ws1,ws2);
        }

        private List<DiferenceWs> PaintDifferences(IXLWorksheet ws1, IXLWorksheet ws2)
        {
            var ws1CountRows = ws1.Rows().Count();
            var ws2CountRows = ws2.Rows().Count();
            var ws1CountCols = ws1.Columns().Count();
            var ws2CountCols = ws2.Columns().Count();
            List<DiferenceWs> diffList = new();

            if (ws1CountCols != ws2CountCols && ws1CountRows != ws2CountRows)
                throw new Exception($"number cols/rows are different for worksheet target {ws1.Name} and worksheet source {ws2.Name} , table : "+TableName);

            for(int irow = 1; irow <= ws1CountRows; irow++)
            {
                for (int icol = 1; icol <= ws1CountCols; icol++)
                {
                    if ((ws1.Cell(irow, icol).GetValue<string>() != ws2.Cell(irow, icol).GetValue<string>()))
                    {
                        ws1.Cell(irow, icol).Style.Fill.BackgroundColor = XLColor.OrangePeel;

                        diffList.Add(new DiferenceWs(irow, icol, ws1.Cell(irow, icol).GetValue<string>(), ws2.Cell(irow, icol).GetValue<string>(), TableName));
                    }
                }
            }
            return diffList;
        }
    }
}
