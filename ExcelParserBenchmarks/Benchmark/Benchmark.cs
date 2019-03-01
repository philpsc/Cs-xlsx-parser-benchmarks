using System.Diagnostics;
using ExcelParserBenchmarks.Enums;
using ExcelParserBenchmarks.Excel_Libraries;

namespace ExcelParserBenchmarks
{
    public class Benchmark 
    {
        private ParserBase Parser;

        public Benchmark(ParserBase parser)
        {
            this.Parser = parser;
        }

        public string Run(Operation operation)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();


            if (operation is Operation.Write)
                Parser.WriteToXlsx();
            else
                Parser.ReadFromXlsx();
            

            stopwatch.Stop();
            var timeSpan = stopwatch.Elapsed.TotalSeconds;
            string elapsedTime = timeSpan.ToString("ss[.ff]");

            return Parser.GetType() + ": " + operation.ToString() + " benchmark completed in: " + timeSpan + "seconds";
        }        
    }
}
