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
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        #region PageMargins

        //For assert this test you need to open html docs and they shouldn't have negative left margins //ToDo: Need to add gold assert tests
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

        #endregion

        //ToDo: Need to add asserts
        #region HtmlOfficeMathOutputMode

        [Test]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Text)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.Text)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text)]
        public void ExportToHtmlUsingImage(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            Document doc = new Document();
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = outputMode;

            Save(doc, "HtmlSaveOptions.ExportToHtmlUsingImage" + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        #endregion

        //ToDo: Add asserts
        #region ExportTextBoxAsSvg

        [Test]
        [TestCase(SaveFormat.Html, true, Description = "TextBox as svg (html)")]
        [TestCase(SaveFormat.Html, false, Description = "TextBox as img (html)")]
        [TestCase(SaveFormat.Mhtml, true, Description = "TextBox as svg (mhtml)")]
        [TestCase(SaveFormat.Mhtml, false, Description = "TextBox as img (mhtml)")]
        [TestCase(SaveFormat.Epub, true, Description = "TextBox as svg (epub)")]
        [TestCase(SaveFormat.Epub, false, Description = "TextBox as img (epub)")]
        public void ExportTextBoxAsSvg(SaveFormat saveFormat, bool textBoxAsSvg)
        {
            Document doc = new Document();

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportTextBoxAsSvg = textBoxAsSvg;

            Save(doc, "HtmlSaveOptions.ExportTextBoxAsSvg" + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        #endregion

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
