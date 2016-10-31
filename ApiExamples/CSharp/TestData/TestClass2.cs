using System.Drawing;
using System.IO;

namespace ApiExamples.TestData
{
    public class TestClass2
    {
        public TestClass2(Stream stream)
        {
            this.Image = stream;
        }

        public TestClass2(Image imageObject)
        {
            this.ImageObject = imageObject;
        }

        public TestClass2(byte[] imageBytes)
        {
            this.ImageBytes = imageBytes;
        }

        public TestClass2(string uriToImage)
        {
            this.UriToImage = uriToImage;
        }

        public Stream Image { get; set; }

        public Image ImageObject { get; set; }

        public byte[] ImageBytes { get; set; }

        public string UriToImage { get; set; }
    }
}
