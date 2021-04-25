using System.Collections.ObjectModel;
using System.IO;
using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace TestImportDataFromExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            //TestImportDataFromExcelFile();
            ExportJsonDataToExcelFile();
        }

        private static void ExportJsonDataToExcelFile()
        {
            //var oldExcelPath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelDataFiles", "Book1.xlsx");
            var newExcelPath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelDataFiles", "Book2.xlsx");
            var jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelDataFiles", "output.json");
            var jsonDataList = new List<CLSExcelToJson>();
            using (StreamReader sr = new StreamReader(jsonPath))
            {
                var jsondata = sr.ReadToEnd();
                if (jsondata != null)
                {
                    jsonDataList = JsonConvert.DeserializeObject<List<CLSExcelToJson>>(jsondata);
                    using (var excelWB = new XLWorkbook())
                    {
                        //var workSheet = excelWB.Worksheet(1);
                        var workSheet = excelWB.Worksheets.Add("JsonExport");
                        if (jsonDataList.Count() > 0)
                        {
                            var rowCount = 1;
                            foreach (var jsondataForExcel in jsonDataList)
                            {
                                workSheet.Cell(rowCount, 1).Value = jsondataForExcel.Column1;
                                workSheet.Cell(rowCount, 2).Value = jsondataForExcel.Column2;
                                workSheet.Cell(rowCount, 3).Value = jsondataForExcel.Column3;
                                workSheet.Cell(rowCount, 4).Value = "Output from program";
                                rowCount++;
                            }

                        }
                        excelWB.SaveAs(newExcelPath);
                    }
                }
            }
        }

        private static void TestImportDataFromExcelFile()
        {
            try
            {
                var excelDataPath = "ExcelDataFiles";
                var fullPath = Path.Combine(Directory.GetCurrentDirectory(), excelDataPath, "Book1.xlsx");
                var jsonPath = Path.Combine(Directory.GetCurrentDirectory(), excelDataPath, "output.json");
                var excelToJsonList = new List<CLSExcelToJson>();
                using (var excelWB = new XLWorkbook(fullPath))
                {
                    var getExcelDataRow = excelWB.Worksheet(1).RowsUsed();
                    if (getExcelDataRow != null)
                    {
                        foreach (var excelData in getExcelDataRow)
                        {
                            var column1 = excelData.Cell(1).Value;
                            var column2 = excelData.Cell(2).Value;
                            var column3 = excelData.Cell(3).Value;
                            Console.WriteLine($"{column1} {column2} {column3}");
                            excelToJsonList.Add(new CLSExcelToJson
                            {
                                Column1 = column1.ToString(),
                                Column2 = column2.ToString(),
                                Column3 = column3.ToString()
                            });
                        }

                        if (excelToJsonList.Count() > 0)
                        {
                            var dataJson = JsonConvert.SerializeObject(excelToJsonList);
                            Console.WriteLine(dataJson);
                            using (StreamWriter sw = new StreamWriter(jsonPath))
                            {
                                sw.Write(dataJson);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                // TODO
            }
        }
    }

    class CLSExcelToJson
    {
        public string Column1 { get; set; }
        public string Column2 { get; set; }
        public string Column3 { get; set; }
    }
}
