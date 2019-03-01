using System;
using System.IO;
using System.Threading;
using ExcelParserBenchmarks.Excel_Libraries;
using OfficeOpenXml;


namespace ExcelParserBenchmarks.Libraries
{
    /// <summary>
    /// The parser that uses the EPPlus library
    /// </summary>
    public class EpplusParser : ParserBase
    {
        public EpplusParser(int[,] testData, string resultSavePath) : base(testData, resultSavePath) { }
              

        override public void WriteToXlsx()
        {
            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("0");

                for (int i = 1; i < TestData.GetLength(0); i++)
                {
                    for (int j = 1; j < TestData.GetLength(1); j++)
                    {
                        worksheet.Cells[i, j].Value = TestData[i, j];
                    }
                }

                var fileName = ResultSavePath + "EpplusBenchmarkResult.xlsx";
                excelPackage.SaveAs(new FileInfo(fileName));
            }
        }

        override public void ReadFromXlsx()
        {
            var fileInfo = new FileInfo(ResultSavePath + "EpplusBenchmarkResult.xlsx");

            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                var worksheet = excelPackage.Workbook.Worksheets[1];

                int rowCount = worksheet.Dimension.End.Row;
                int columnCount = worksheet.Dimension.End.Column;
                int[,] data = new int[rowCount + 1, columnCount + 1];

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= columnCount; j++)
                    {
                        data[i, j] = Convert.ToInt32(worksheet.Cells[i, j].Value);
                    }
                }
            }
        }
    }
}
