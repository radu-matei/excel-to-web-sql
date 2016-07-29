using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace ExcelToWebSql.Generator
{
    public class ExcelSqlScriptGenerator : IExcelSqlScriptGenerator
    {
        private Application _excelApplication { get; set; }
        private Workbook _workbook { get; set; }

        public ExcelSqlScriptGenerator(string path)
        {
            _excelApplication = new Application();
            _workbook = _excelApplication.Workbooks.Open(path);
        }

        public void GenerateSqlTableScipts()
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

            for(int sheetNumber = 1; sheetNumber <= _workbook.Sheets.Count; sheetNumber++)
            {
                Worksheet sheet = _workbook.Sheets[sheetNumber];
                string sqlCreateTableStatement = "CREATE TABLE " + sheet.Name + "( \n";

                for(int column = 1; column <= sheet.UsedRange.Columns.Count; column++)
                {
                    sqlCreateTableStatement += "\n" + sheet.Cells[1, column].Text + " " + formatDictionary[sheet.Cells[1, column].NumberFormat] + ",";
                }

                sqlCreateTableStatement += "\n)";

                Console.WriteLine(sqlCreateTableStatement);
            }
        }
        public void GenerateSqlInsertScripts()
        {
            for (int sheetNumnber = 1; sheetNumnber <= _workbook.Sheets.Count; sheetNumnber++)
            {
                Worksheet sheet = _workbook.Sheets.Item[sheetNumnber];

                string sqlInsertStatement = "INSERT INTO " + sheet.Name + "( ";
                var columns = GetSheetColumns(sheet);

                for (int i = 0; i < columns.Count - 1; i++)
                    sqlInsertStatement += columns[i] + ", ";

                sqlInsertStatement += columns[columns.Count - 1] + " )" + "\n" + "VALUES ";

                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    sqlInsertStatement += "\n       ( ";

                    for (int column = 1; column <= sheet.UsedRange.Columns.Count - 1; column++)
                    {
                        if (sheet.Cells[row, column].NumberFormat == "General" || sheet.Cells[row, column].NumberFormat == "@")
                            sqlInsertStatement += String.Format("'{0}', ", sheet.Cells[row, column].Value2.ToString());

                        else
                            sqlInsertStatement += String.Format("{0}, ", sheet.Cells[row, column].Value2.ToString());
                    }

                    if (sheet.Cells[row, sheet.UsedRange.Columns.Count].NumberFormat == "General" || sheet.Cells[row, sheet.UsedRange.Columns.Count].NumberFormat == "@")
                        sqlInsertStatement += String.Format("'{0}'", sheet.Cells[row, sheet.UsedRange.Columns.Count].Value2.ToString()) + " )\n      ";

                    else
                        sqlInsertStatement += String.Format("{0}", sheet.Cells[row, sheet.UsedRange.Columns.Count].Value2.ToString()) + " )\n      ";
                }
                
                Console.WriteLine(sqlInsertStatement);
            }
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
            _workbook.Close();
            _excelApplication.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApplication);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);

            GC.Collect();
        }
    }
}
