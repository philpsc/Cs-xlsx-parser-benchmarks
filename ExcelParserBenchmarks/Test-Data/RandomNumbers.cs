using System;

namespace ExcelParserBenchmarks.Test_Data
{
    public class RandomNumbers
    {
        private readonly int rows;
        private readonly int cols;

        public RandomNumbers(int Rows, int Cols)
        {
            // +1: Rows and columns are 1-based in Excel sheets
            rows = Rows +1;
            cols = Cols +1;
        }
        public int[,] ProduceIntMatrix()
        {
            var randomNumbers = new int[rows,cols];
            var random = new Random();
            for (int i = 1; i < cols; i++)
            {
                for (int j = 1; j < cols; j++)
                {
                    randomNumbers[i, j] = random.Next(9);
                }
            }            

            return randomNumbers;
        }
    }
}
