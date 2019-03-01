using ExcelParserBenchmarks.Enums;
using ExcelParserBenchmarks.Libraries;
using ExcelParserBenchmarks.Test_Data;
using System;
using System.Collections.Generic;


namespace ExcelParserBenchmarks
{
    public class BenchmarkLauncher
    {
        public void Start()
        {
            // Pfad, wo generierte .xlsx-Dateien gespeichert werden
            var resultSavePath = "..\\..\\Test-Data\\";

            // Zufallszahlen: Zweidimensionales Array mit 50.000 Zeilen, 100 Spalten 
            var testData = new RandomNumbers(50000, 100).ProduceIntMatrix();

            var epplusParser = new EpplusParser(testData, resultSavePath);
            var epplusBenchmark = new Benchmark(epplusParser);

            var npoiParser = new NpoiParser(testData, resultSavePath);
            var npoiBenchmark = new Benchmark(npoiParser);

            var closedXmlParser = new ClosedXmlParser(testData, resultSavePath);
            var closedXmlBenchmark = new Benchmark(new ClosedXmlParser(testData, resultSavePath));

            var benchmarks = new List <Benchmark>(){epplusBenchmark, npoiBenchmark, closedXmlBenchmark};
            
            foreach(var benchmark in benchmarks)
            {
                Console.WriteLine(benchmark.Run(Operation.Write));
                Console.WriteLine(benchmark.Run(Operation.Read));
            }
        }
    }
}

