using System;

namespace ApiExamples.TestData
{
    public class TestClass3
    {
        public TestClass3(int firstNumber, double secondNumber, int thirdNumber, DateTime date)
        {
            this.FirstNumber = firstNumber;
            this.SecondNumber = secondNumber;
            this.ThirdNumber = thirdNumber;
            this.Date = date;
        }

        public int FirstNumber { get; set; }

        public double SecondNumber { get; set; }

        public int ThirdNumber { get; set; }

        public DateTime Date { get; set; }
    }
}
