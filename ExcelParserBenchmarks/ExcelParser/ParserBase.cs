namespace ExcelParserBenchmarks.Excel_Libraries
{
    public abstract class ParserBase
    {
        protected int[,] TestData { get; }
        protected string ResultSavePath { get; }

        public ParserBase(int[,] testData, string resultSavePath)
        {
            this.TestData = testData;
            this.ResultSavePath = resultSavePath;
        }

        public abstract void ReadFromXlsx();

        public abstract void WriteToXlsx();

    }
}
