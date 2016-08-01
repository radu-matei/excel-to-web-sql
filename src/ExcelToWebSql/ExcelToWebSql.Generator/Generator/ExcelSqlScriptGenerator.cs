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
                AppendColumnNames(sheet, ref sqlCreateTableStatement);

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
                var columns = GetSheetColumns(sheet);

                for (int i = 0; i < columns.Count - 1; i++)
                    sqlInsertStatement += columns[i] + ", ";

                sqlInsertStatement += columns[columns.Count - 1] + " )" + "\n" + "VALUES ";

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
                sqlStatementBuilder.Append("\n       ( ");

                for (int column = 1; column <= sheet.UsedRange.Columns.Count - 1; column++)
                {
                    if (sheet.Cells[row, column].NumberFormat == "General"
                       || sheet.Cells[row, column].NumberFormat == "@")

                        sqlStatementBuilder.AppendFormat("'{0}', ", sheet.Cells[row, column].Value2.ToString());

                    else
                        sqlStatementBuilder.AppendFormat("{0}, ", sheet.Cells[row, column].Value2.ToString());
                }

                if (sheet.Cells[row, sheet.UsedRange.Columns.Count].NumberFormat == "General"
                   || sheet.Cells[row, sheet.UsedRange.Columns.Count].NumberFormat == "@")

                    sqlStatementBuilder.AppendFormat("'{0}'", sheet.Cells[row, sheet.UsedRange.Columns.Count].Value2.ToString())
                                       .AppendFormat(" )\n      ");

                else
                    sqlStatementBuilder.AppendFormat("{0}", sheet.Cells[row, sheet.UsedRange.Columns.Count].Value2.ToString())
                                       .AppendFormat(" )\n      ");
            }
            sqlInsertStatement = sqlStatementBuilder.ToString();
        }

        private void AppendColumnNames(Worksheet sheet, ref string sqlCreateTableStatement)
        {
            var formatDictionary = new Dictionary<string, string>()
            {
                {"@", "nvarchar(50)" },
                {"General", "nvarchar(50)" },
                {"0", "int" },
                {"0,0", "float" },
                {"0,00", "float" },
                {"0,000", "float" }
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
