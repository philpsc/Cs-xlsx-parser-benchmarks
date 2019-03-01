using System;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using ExcelParserBenchmarks.Excel_Libraries;

namespace ExcelParserBenchmarks.Libraries
{
    /// <summary>
    /// The parser that uses the ClosedXML library.
    /// </summary>
    public class ClosedXmlParser : ParserBase
    {
        public ClosedXmlParser(int[,] testData, string resultSavePath) : base(testData, resultSavePath) {}
        
        override public void WriteToXlsx()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("0");

                for (int i = 1; i < TestData.GetLength(0); i++)
                {
                    for (int j = 1; j < TestData.GetLength(1); j++)
                    {
                        worksheet.Cell(i, j).Value = TestData[i, j];
                    }
                }

                workbook.SaveAs(ResultSavePath + "ClosedXMLBenchmarkResult.xlsx");
            }
        }

        override public void ReadFromXlsx()
        {
            using (var workbook = new XLWorkbook(ResultSavePath + "ClosedXMLBenchmarkResult.xlsx"))
            {
                var worksheet = workbook.Worksheets.Worksheet("0");

                int rowCount = worksheet.Rows().Count();
                int columnCount = worksheet.Columns().Count();
                int[,] data = new int[rowCount+1, columnCount+1];

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= columnCount; j++)
                    {
                        data[i, j] = Convert.ToInt32(worksheet.Cell(i,j).GetDouble());
                    }
                }
            }
        }
    }
}
