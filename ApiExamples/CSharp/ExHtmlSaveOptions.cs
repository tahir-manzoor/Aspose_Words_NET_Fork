// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;

using NUnit.Framework;

namespace ApiExamples
{
    //ToDo: Need to add gold assert tests
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        //For assert this test you need to open html docs and they shouldn't have negative left margins
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = saveFormat;
            saveOptions.ExportPageMargins = true;

            Save(doc, "HtmlSaveOptions.ExportPageMargins." + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportToHtmlUsingImage(SaveFormat saveFormat)
        {
            Document doc = new Document();
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = HtmlOfficeMathOutputMode.Image;

            Save(doc, "HtmlSaveOptions.ExportToHtmlUsingImage" + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportToHtmlUsingMathMl(SaveFormat saveFormat)
        {
            Document doc = new Document();

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML;

            Save(doc, "HtmlSaveOptions.ExportToHtmlUsingImage" + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportToHtmlUsingText(SaveFormat saveFormat)
        {
            Document doc = new Document();

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = HtmlOfficeMathOutputMode.Text;

            Save(doc, "HtmlSaveOptions.ExportToHtmlUsingImage" + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        private static Document Save(Document inputDoc, string outputDocPath, SaveFormat saveFormat, SaveOptions saveOptions)
        {
            switch (saveFormat)
            {
                case SaveFormat.Html:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return inputDoc;
                case SaveFormat.Mhtml:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return inputDoc;
                case SaveFormat.Epub:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions); //There is draw images bug with epub. Need write to NSezganov
                    return inputDoc;
            }

            return null;
        }
    }
}
