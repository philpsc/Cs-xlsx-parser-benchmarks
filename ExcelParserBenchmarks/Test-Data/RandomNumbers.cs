using System;

namespace ExcelParserBenchmarks.Test_Data
{
    public class RandomNumbers
    {
        private int Rows { get; }
        private int Cols { get; }

        public RandomNumbers(int Rows, int Cols)
        {
            // +1: Rows and columns are 1-based in Excel sheets
            this.Rows = Rows +1;
            this.Cols = Cols +1;
        }
        public int[,] ProduceIntMatrix()
        {
            var randomNumbers = new int[Rows,Cols];
            var random = new Random();
            for (int i = 1; i < Rows; i++)
            {
                for (int j = 1; j < Cols; j++)
                {
                    randomNumbers[i, j] = random.Next(9);
                }
            }            

            return randomNumbers;
        }
    }
}
