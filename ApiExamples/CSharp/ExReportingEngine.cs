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
using System.Runtime.CompilerServices;
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
        private readonly string _doc = MyDir + "ReportingEngine.TestDataTable.docx";

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
            Document doc = new Document(MyDir + "ReportingEngine.TestDataTable.docx");

            DataSet ds = TestTables.AddClientsTestData();
            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + "ReportingEngine.TestDataTable Out.docx");
        }

        [Test]
        public void NestedDataTableTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestNestedDataTable.docx");

            DataSet ds = TestTables.AddClientsTestData();
            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + "ReportingEngine.TestNestedDataTable Out.docx");
        }

        [Test]
        public void ChartTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestChart.docx");

            DataSet ds = TestTables.AddClientsTestData();

            //var sumManager1 = ds.Contracts.Where(c => c.ManagerId == "1").Sum(c => c.Price);
            //var sumManager2 = ds.Contracts.Where(c => c.ManagerId == "2").Sum(c => c.Price);
            //var sumManager3 = ds.Contracts.Where(c => c.ManagerId == "3").Sum(c => c.Price);
            var sumManager3 = ds.Contracts.Count();

            BuildReport(doc, ds.Managers, "managers");

            doc.Save(MyDir + "ReportingEngine.TestChart Out.docx");
        }

        [Test]
        public void BubbleChartTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestBubbleChart.docx");
            DataSet ds = TestTables.AddClientsTestData();

            BuildReport(doc, ds.Managers, "managers");
            
            doc.Save(MyDir + "ReportingEngine.TestBubbleChart Out.docx");
        }

        [Test]
        public void IndexOf()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestIndexOf.docx");

            DataSet ds = TestTables.AddClientsTestData();
            
            BuildReport(doc, ds, "ds");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("The names are: Name 1, Name 2, Name 3\f", doc.GetText());
        }

        [Test]
        public void IfElse()
        {
            Document doc = new Document(MyDir + "ReportingEngine.IfElse.docx");

            DataSet ds = TestTables.AddClientsTestData();

            BuildReport(doc, ds.Managers, "m");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("You have chosen 3 item(s).\f", doc.GetText());
        }

        [Test]
        public void IfElseWithoutData()
        {
            Document doc = new Document(MyDir + "ReportingEngine.IfElse.docx");

            DataSet ds = new DataSet();

            BuildReport(doc, ds.Managers, "m");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("You have chosen no items.\f", doc.GetText());
        }

        //ToDo:Need to assert with doc without contaxual object member access
        [Test]
        public void ContextualObjectMemberAccess()
        {
            Document doc = new Document(MyDir + "ReportingEngine.ContextualObjectMemberAccess.docx");

            DataSet ds = TestTables.AddClientsTestData();

            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + "ReportingEngine.ContextualObjectMemberAccess Out.docx");
        }

        //Add byte, Document, Uri
        [Test]
        public void InsertDocumentDinamically()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<doc [src.DocumentByByte]>>");

            //string url = "http://www.aspose.com/demos/.net-components/aspose.words/csharp/general/Common/Documents/DinnerInvitationDemo.doc";
            byte[] bytes = File.ReadAllBytes(MyDir + "ReportingEngine.IfElse.docx");

            //TestClass4 docStream = new TestClass4(new FileStream(this._doc, FileMode.Open, FileAccess.Read));
            //TestClass4 docByDoc = new TestClass4(new Document(MyDir + "ReportingEngine.IfElse.docx"));
            //TestClass4 docByUri = new TestClass4(url);
            TestClass4 docByByte = new TestClass4(bytes);

            BuildReport(doc, docByByte, "src", ReportBuildOptions.None);

            doc.Save(MyDir + "ReportingEngine.InsertDocumentDinamically Out.docx");
        }

        //Add byte, Document, Uri
        [Test]
        public void InsertImageDinamically()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image]>>");

            //string url = "http://www.aspose.com/demos/.net-components/aspose.words/csharp/general/Common/Documents/DinnerInvitationDemo.doc";
            byte[] bytes = File.ReadAllBytes(MyDir + "ReportingEngine.IfElse.docx");

            //TestClass2 docStream = new TestClass2(new FileStream(this._doc, FileMode.Open, FileAccess.Read));
            //TestClass2 docByDoc = new TestClass2(new Document(MyDir + "ReportingEngine.IfElse.docx"));
            //TestClass2 docByUri = new TestClass2(url);
            TestClass2 docByByte = new TestClass2(bytes);

            BuildReport(doc, docByByte, "src", ReportBuildOptions.None);

            doc.Save(MyDir + "ReportingEngine.InsertDocumentDinamically Out.docx");
        }

        [Test]
        public void WithoutKnownType()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[new DateTime()]:”dd.MM.yyyy”>>");
           
            ReportingEngine engine = new ReportingEngine();
            Assert.That(() => engine.BuildReport(doc, ""), Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void WorkWithKnownTypes()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[DateTime.Now]:”dd.MM.yyyy”>>");
            builder.Writeln("<<[DateTime.Now]:”dd”>>");
            builder.Writeln("<<[DateTime.Now]:”MM”>>");
            builder.Writeln("<<[DateTime.Now]:”yyyy”>>");

            builder.Writeln("<<[DateTime.Now.Month]>>");

            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(DateTime));
            engine.BuildReport(doc, "");
            
            doc.Save(MyDir + "ReportingEngine.KnownTypes Out.docx");
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