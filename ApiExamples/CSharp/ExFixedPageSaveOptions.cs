// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Text;

using Aspose.Words;
using Aspose.Words.Saving;

using NUnit.Framework;

using System.Collections.Generic;

namespace ApiExamples
{
    [TestFixture]
    internal class ExFixedPageSaveOptions : ApiExampleBase
    {
        public IEnumerable<TestCaseData> CountEqualsZeroAndHouseGrossIsGreaterTestCases
        {
            get
            {
                yield return new TestCaseData(new XamlFixedSaveOptions());
                yield return new TestCaseData(new XpsSaveOptions());
                yield return new TestCaseData(new SwfSaveOptions());
            }
        }

        public IEnumerable<TestCaseData> MetafileRenderingOptions
        {
            get
            {
                yield return new TestCaseData(new XamlFixedSaveOptions(), MetafileRenderingMode.Bitmap);
                yield return new TestCaseData(new XamlFixedSaveOptions(), MetafileRenderingMode.Vector);
                yield return new TestCaseData(new XamlFixedSaveOptions(), MetafileRenderingMode.VectorWithFallback);
                yield return new TestCaseData(new XpsSaveOptions(), MetafileRenderingMode.Bitmap);
                yield return new TestCaseData(new XpsSaveOptions(), MetafileRenderingMode.Vector);
                yield return new TestCaseData(new XpsSaveOptions(), MetafileRenderingMode.VectorWithFallback);
                yield return new TestCaseData(new SwfSaveOptions(),MetafileRenderingMode.Bitmap);
                yield return new TestCaseData(new SwfSaveOptions(), MetafileRenderingMode.Vector);
                yield return new TestCaseData(new SwfSaveOptions(), MetafileRenderingMode.VectorWithFallback);
            }
        }

        //ToDo: Need to add asserts
        [Test]
        [TestCaseSource(nameof(CountEqualsZeroAndHouseGrossIsGreaterTestCases))]
        public void JpegQualityDefaultValue(FixedPageSaveOptions objectSaveOptions)
        {
            FixedPageSaveOptions saveOptions = objectSaveOptions;
            Assert.AreEqual(95, saveOptions.JpegQuality);
        }

        [Test]
        [TestCaseSource(nameof(MetafileRenderingOptions))]
        public void MetafileRendering(FixedPageSaveOptions objectSaveOptions, MetafileRenderingMode metafileRendering)
        {
            Document doc = new Document();

            FixedPageSaveOptions saveOptions = objectSaveOptions;
            saveOptions.MetafileRenderingOptions.RenderingMode = metafileRendering;

            doc.Save(MyDir + @"\Artifacts\MetafileRendering.swf", saveOptions);
        }
    }
}
