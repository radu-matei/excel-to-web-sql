using Microsoft.Office.Interop.Excel;
using System.Dynamic;
using System;
using System.Collections.Generic;
using ExcelToWebSql.Generator.Storage;

namespace ExcelToWebSql.Generator
{
    public class ExcelConverter : IExcelConverter
    {
        private Application _excelApplication { get; set; }
        private Workbook _workbook { get; set; }
        private IStorage _storage { get; set; }


        public ExcelConverter()
        {
            _excelApplication = new Application();
            _storage = new FileStorage();
        }

        public void ConvertToJson(string sourceFilePath, string outputFilePath)
        {
            var jsonSerializer = new JsonSerializer();
            
            foreach(var sheetList in ConvertToObject(sourceFilePath))
            {
                var fileName = outputFilePath + sheetList.Key + ".json";
                _storage.SaveDocument(jsonSerializer.Serialize(sheetList.Value), fileName);
            }
        }

        public void ConvertToXml(string sourceFilePath,string outputFilePath)
        {
            var xmlSerializer = new XmlSerializer();

            foreach (var sheetList in ConvertToObject(sourceFilePath))
            {
                var fileName = outputFilePath + sheetList.Key + ".xml";
                _storage.SaveDocument(xmlSerializer.Serialize(sheetList.Value), fileName);
            }
        }

        private Dictionary<string, List<ExpandoObject>> ConvertToObject(string sourceFilePath)
        {
            _workbook = _excelApplication.Workbooks.Open(sourceFilePath);
            var workbookDictionary = new Dictionary<string, List<ExpandoObject>>();

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
                workbookDictionary.Add(sheet.Name, dynamicSheetList);
            }
            _workbook.Close();
            return workbookDictionary;
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
