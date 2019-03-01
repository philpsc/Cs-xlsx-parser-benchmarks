namespace ExcelParserBenchmarks.Excel_Libraries
{
    public abstract class ParserBase
    {
        protected int[,] testData;
        protected string resultSavePath;

        public ParserBase(int[,] testData, string resultSavePath)
        {
            this.testData = testData;
            this.resultSavePath = resultSavePath;
        }

        public abstract void ReadFromXlsx();

        public abstract void WriteToXlsx();

    }
}
