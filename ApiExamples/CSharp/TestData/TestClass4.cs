using System.IO;
using Aspose.Words;

namespace ApiExamples.TestData
{
    public class TestClass4
    {
        public TestClass4(Document doc)
        {
            this.Document = doc;
        }

        public TestClass4(Stream stream)
        {
            this.DocumentByStream = stream;
        }

        public TestClass4(byte[] byteDoc)
        {
            this.DocumentByByte = byteDoc;
        }

        public TestClass4(string uriToDoc)
        {
            this.DocumentByUri = uriToDoc;
        }

        public Document Document { get; set; }

        public Stream DocumentByStream { get; set; }
        
        public byte[] DocumentByByte { get; set; }

        public string DocumentByUri { get; set; }

    }
}
