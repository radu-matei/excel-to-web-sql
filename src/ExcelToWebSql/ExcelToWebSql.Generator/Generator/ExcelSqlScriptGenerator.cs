using ExcelToWebSql.Generator.Storage;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToWebSql.Generator
{
    public class ExcelSqlScriptGenerator : IExcelSqlScriptGenerator
    {
        private Application _excelApplication { get; set; }
        private Workbook _workbook { get; set; }
        private IStorage _storage { get; set; }

        public ExcelSqlScriptGenerator()
        {
            _excelApplication = new Application();
            _storage = new FileStorage();
        }

        public void GenerateSqlTableScipts(string sourceFilePath, string outputFilePath)
        {
            _workbook = _excelApplication.Workbooks.Open(sourceFilePath);

            for(int sheetNumber = 1; sheetNumber <= _workbook.Sheets.Count; sheetNumber++)
            {
                Worksheet sheet = _workbook.Sheets[sheetNumber];
                string sqlCreateTableStatement = "CREATE TABLE " + sheet.Name + "( \n";
                AddSqlColumnMapping(sheet, ref sqlCreateTableStatement);

                var fileName = outputFilePath + "CREATE " + sheet.Name + ".sql";
                _storage.SaveDocument(sqlCreateTableStatement, fileName);
            }
            _workbook.Close();
        }
        public void GenerateSqlInsertScripts(string sourceFilePath, string outputFilePath)
        {
            _workbook = _excelApplication.Workbooks.Open(sourceFilePath);

            for (int sheetNumnber = 1; sheetNumnber <= _workbook.Sheets.Count; sheetNumnber++)
            {
                Worksheet sheet = _workbook.Sheets.Item[sheetNumnber];
                string sqlInsertStatement = "INSERT INTO " + sheet.Name + "( ";

                AddSqlColumnNames(sheet, ref sqlInsertStatement);
                sqlInsertStatement += "\n" + "VALUES ";
                AppendInsertValues(sheet, ref sqlInsertStatement);
                
                var fileName = outputFilePath + "INSERT " + sheet.Name + ".sql";
                _storage.SaveDocument(sqlInsertStatement, fileName);
            }
            _workbook.Close();
        }

        private void AppendInsertValues(Worksheet sheet, ref string sqlInsertStatement)
        {
            StringBuilder sqlStatementBuilder = new StringBuilder(sqlInsertStatement);
            for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
            {
                AppendSqlRowInsert(sheet, row, sqlStatementBuilder);
                if (row != sheet.UsedRange.Rows.Count)
                    sqlStatementBuilder.AppendFormat(",");
            }

            sqlInsertStatement = sqlStatementBuilder.ToString();
        }

        private void AppendSqlRowInsert(Worksheet sheet, int rowNumber, StringBuilder sqlStatementBuilder)
        {
            sqlStatementBuilder.Append("\n       ( ");

            for (int column = 1; column <= sheet.UsedRange.Columns.Count; column++)
            {
                if (sheet.Cells[rowNumber, column].NumberFormat == "General"
                                       || sheet.Cells[rowNumber, column].NumberFormat == "@")

                    sqlStatementBuilder.AppendFormat("'{0}'", sheet.Cells[rowNumber, column].Text.ToString());

                else
                    sqlStatementBuilder.AppendFormat("{0}", sheet.Cells[rowNumber, column].Text.ToString());

                if (column != sheet.UsedRange.Columns.Count)
                    sqlStatementBuilder.AppendFormat(", ");
            }
            sqlStatementBuilder.AppendFormat(")");
        }

        private void AddSqlColumnMapping(Worksheet sheet, ref string sqlCreateTableStatement)
        {
            var formatDictionary = new Dictionary<string, string>()
            {
                {"@", "nvarchar(50)" },
                {"General", "nvarchar(50)" },
                {"0", "int" },
                {"0.0", "decimal" },
                {"0.00", "decimal" },
                {"0.000", "decimal" }
            };

            StringBuilder sqlStatementBuilder = new StringBuilder(sqlCreateTableStatement);

            for (int column = 1; column <= sheet.UsedRange.Columns.Count; column++)
            {
                sqlStatementBuilder.Append("\n" + sheet.Cells[1, column].Text + " ")
                                   .Append(formatDictionary[sheet.Cells[1, column].NumberFormat])
                                   .Append(",");
            }
            sqlStatementBuilder.Append("\n)");

            sqlCreateTableStatement = sqlStatementBuilder.ToString();
        }

        private void AddSqlColumnNames(Worksheet sheet, ref string sqlStatement)
        {
            StringBuilder sqlStatementBuilder = new StringBuilder(sqlStatement);
            var columns = GetSheetColumns(sheet);

            for (int i = 0; i < columns.Count - 1; i++)
                sqlStatementBuilder.AppendFormat(columns[i] + ", ");

            sqlStatementBuilder.AppendFormat(columns[columns.Count - 1])
                               .Append(" )");

            sqlStatement = sqlStatementBuilder.ToString();
        }

        private List<string> GetSheetColumns(Worksheet sheet)
        {
            var columns = new List<string>();

            for (int i = 1; i <= sheet.UsedRange.Columns.Count; i++)
            {
                columns.Add(sheet.Cells[1, i].Text);
            }

            return columns;
        }

        public void Dispose()
        {
            _excelApplication.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApplication);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);

            GC.Collect();
        }
    }
}
