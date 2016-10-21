// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.IO;
using System.Linq;
using ApiExamples.TestData;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using NUnit.Framework;
using DataSet = ApiExamples.TestData.DataSet;

namespace ApiExamples
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        private readonly string _image = MyDir + "Test_636_852.gif";

        [Test]
        public void SimpleCase()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<[s.Name]>> says: <<[s.Message]>>");

            TestClass1 sender = new TestClass1("LINQ Reporting Engine", "Hello World");

            BuildReport(doc, sender, "s", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("LINQ Reporting Engine says: Hello World\f", doc.GetText());
        }

        [Test]
        public void StringFormat()
        {
            Document doc =
                DocumentHelper.CreateTemplateDocumentForReportingEngine(
                    "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

            TestClass1 sender = new TestClass1("LINQ Reporting Engine", "hello world");
            BuildReport(doc, sender, "s");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f", doc.GetText());
        }

        [Test]
        public void NumberFormat()
        {
            Document doc =
                DocumentHelper.CreateTemplateDocumentForReportingEngine(
                    "<<[s.FirstNumber]:alphabetic>> : <<[s.SecondNumber]:roman:lower>>, <<[s.ThirdNumber]:ordinal>>, <<[s.FirstNumber]:ordinalText:upper>>" +
                    ", <<[s.SecondNumber]:cardinal>>, <<[s.ThirdNumber]:hex>>, <<[s.ThirdNumber]:arabicDash>>, <<[s.Date]:\"MMMM\":lower>>");

            TestClass3 sender = new TestClass3(1, 2.2, 200, DateTime.Parse("10.09.2016 10:00:00"));
            BuildReport(doc, sender, "s");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("A : ii, 200th, FIRST, Two, C8, - 200 -, сентябрь\f", doc.GetText());
        }

        [Test]
        public void DataTableTest()
        {
            Document doc = new Document(MyDir + "TestDataTable.docx");

            DataSet ds = TestTables.AddClientsTestData();
            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + "TestDataTable Out.docx");
        }

        [Test]
        public void StretchImagefitHeight()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image] -fitHeight>>", ShapeType.TextBox);

            TestClass2 imageStream = new TestClass2(new FileStream(this._image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that width is keeped and height is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        public void StretchImagefitWidth()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image] -fitWidth>>", ShapeType.TextBox);

            TestClass2 imageStream = new TestClass2(new FileStream(this._image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox and 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that height is keeped and width is changed
                Assert.AreNotEqual(431.5, shape.Width);
                Assert.AreEqual(346.35, shape.Height);
            }

            dstStream.Dispose();
        }

        [Test]
        public void StretchImagefitSize()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image] -fitSize>>", ShapeType.TextBox);

            TestClass2 imageStream = new TestClass2(new FileStream(this._image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that height is changed and width is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreNotEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        public void WithoutMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder, new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            //Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
            Assert.That(() => BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.None), Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void WithMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder, new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.AllowMissingMembers);

            //Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
            Assert.AreEqual(
            ControlChar.ParagraphBreak + ControlChar.ParagraphBreak + ControlChar.SectionBreak,
            builder.Document.GetText());
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName, ReportBuildOptions reportBuildOptions)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.Options = reportBuildOptions;

            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object[] dataSource, string[] dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object dataSource)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource);
        }
    }
}