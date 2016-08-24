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
    using System.Data;

    [TestFixture]
    internal class ExImageSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseGdiEmfRenderer()
        {
            //ExStart
            //ExFor:ImageSaveOptions.UseGdiEmfRenderer
            //ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf);
            saveOptions.UseGdiEmfRenderer = false;
            //ExEnd
        }

        [Test]
        public void SaveIntoGif()
        {
            // Test case 1: Assert that saving into gif format correct
            //Test case 2: Assert that pageindex that we try to save are correct
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Gif);
            saveOptions.PageIndex = 1;

            doc.Save(MyDir + @"\Artifacts\SaveOptions.MyraidPro Out.gif", saveOptions);
        }
    }
}
