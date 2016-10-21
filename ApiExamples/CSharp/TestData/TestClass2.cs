using System.IO;

namespace ApiExamples.TestData
{
    public class TestClass2
    {
        public TestClass2(Stream stream)
        {
            this.Image = stream;
        }

        public Stream Image { get; set; }
    }
}
