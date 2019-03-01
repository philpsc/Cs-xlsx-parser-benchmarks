using NPOI.XSSF.UserModel;
using System.IO;
using ExcelParserBenchmarks.Excel_Libraries;
using System;

namespace ExcelParserBenchmarks.Libraries
{
    /// <summary>
    /// The parser that uses the NPOI library.
    /// </summary>
    public class NpoiParser : ParserBase
    {
        public NpoiParser(int[,] testData, string resultSavePath) : base(testData, resultSavePath){}

        override public void WriteToXlsx()
        {
            var workbook = new XSSFWorkbook(); // Doesn't implement IDisposable; can't be used with "using"-block :(
            var worksheet = workbook.CreateSheet("0");

            for (int i = 1; i < TestData.GetLength(0); i++)
            {
                var row = worksheet.CreateRow(i);
                for (int j = 1; j < TestData.GetLength(1); j++)
                {
                    row.CreateCell(j).SetCellValue(TestData[i, j]);
                }
            }

            using (var fileStream = File.Create(ResultSavePath + "NpoiBenchmarkResult.xlsx"))
            {
                workbook.Write(fileStream);
            }
            workbook.Close();
        }

        override public void ReadFromXlsx()
        {
            XSSFWorkbook workbook;
            var filePath = ResultSavePath + "NpoiBenchmarkResult.xlsx";
            using (var fileStream = new FileStream(@filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            var worksheet = workbook.GetSheet("0");

            var rowCount = worksheet.LastRowNum;
            var columnCount = worksheet.GetRow(1).LastCellNum-1;
            int[,] data = new int[rowCount + 1, columnCount + 1];

            for (int i = 1; i <= rowCount; i++)
            {
                var row = worksheet.GetRow(i);
                for (int j = 1; j <= columnCount; j++)
                {
                    data[i, j] = Convert.ToInt32(row.GetCell(j).NumericCellValue);
                }
            }
            workbook.Close();
        }        
    }
}
