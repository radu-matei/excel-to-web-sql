using Microsoft.Office.Interop.Excel;
using System.Dynamic;
using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ExcelToWebSql.Generator
{
    public class ExcelConverter : IExcelConverter
    {
        private Application _excelApplication { get; set; }
        private Workbook _workbook { get; set; }
        private enum ExportTye { Json, Xml }


        public string Path { get; set; }

        public ExcelConverter(string path)
        {
            _excelApplication = new Application();
            _workbook = _excelApplication.Workbooks.Open(path);

            Path = path;
        }

        public void ConvertToJson(string outputPath)
        {
            Convert(outputPath, ExportTye.Json);
        }

        public void ConvertToXml(string outputPath)
        {
            Convert(outputPath, ExportTye.Xml);
        }

        private void Convert(string outputPath, ExportTye type)
        {
            for (int sheetNumber = 1; sheetNumber <= _workbook.Sheets.Count; sheetNumber++)
            {
                var sheet = _workbook.Sheets[sheetNumber];
                var dynamicSheetList = new List<ExpandoObject>();

                for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
                {
                    var dynamicRowObject = new ExpandoObject();

                    for (int column = 1; column <= sheet.UsedRange.Columns.Count; column++)
                    {
                        var value = sheet.Cells[row, column].Value2 as object;
                        DynamicExtensions.AddProperty(dynamicRowObject, sheet.Cells[1, column].Text, value);
                    }
                    dynamicSheetList.Add(dynamicRowObject);
                }
                foreach (var dynamicObject in dynamicSheetList)
                {
                    dynamicObject.Print();
                }
            }
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
