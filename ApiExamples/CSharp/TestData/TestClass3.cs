using System;
using Aspose.Pdf;

namespace ApiExamples.TestData
{
    public class TestClass3
    {
        public TestClass3(int firstNumber, double secondNumber, int thirdNumber, bool logical, DateTime date)
        {
            this.FirstNumber = firstNumber;
            this.SecondNumber = secondNumber;
            this.ThirdNumber = thirdNumber;
            this.Logical = logical;
            this.Date = date;
        }

        public int FirstNumber { get; set; }

        public double SecondNumber { get; set; }

        public int ThirdNumber { get; set; }

        public bool Logical { get; set; }

        public DateTime Date { get; set; }

        public int Sum(int firstNumber, int secondNumber)
        {
            int result = firstNumber + secondNumber;

            return result;
        }

        public bool Pass(bool res)
        {
            return res;
        }

        public int? Test(int? res)
        {
            return res;
        }
    }
}
