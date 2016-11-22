using System.Drawing;
using System.IO;

namespace ApiExamples.TestData
{
    public class TestClass2
    {
        public TestClass2(Stream stream)
        {
            this.Stream = stream;
        }

        public TestClass2(Image imageObject)
        {
            this.Image = imageObject;
        }

        public TestClass2(byte[] imageBytes)
        {
            this.Bytes = imageBytes;
        }

        public TestClass2(string uriToImage)
        {
            this.Uri = uriToImage;
        }

        public Stream Stream { get; set; }

        public Image Image { get; set; }

        public byte[] Bytes { get; set; }

        public string Uri { get; set; }
    }
}
