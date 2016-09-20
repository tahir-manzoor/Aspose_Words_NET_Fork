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
using System.IO;
using Aspose.Words.Rendering;

namespace ApiExamples
{
    [TestFixture]
    internal class ExFixedPageSaveOptions : ApiExampleBase
    {
        public static IEnumerable<TestCaseData> JpegQualityData
        {
            get
            {
                yield return new TestCaseData(new HtmlFixedSaveOptions());
                yield return new TestCaseData(new ImageSaveOptions(SaveFormat.Bmp));
                yield return new TestCaseData(new PdfSaveOptions());
                yield return new TestCaseData(new PsSaveOptions());
                yield return new TestCaseData(new SvgSaveOptions());
                yield return new TestCaseData(new XamlFixedSaveOptions());
                yield return new TestCaseData(new XpsSaveOptions());
                yield return new TestCaseData(new SwfSaveOptions());
            }
        }

        public static IEnumerable<TestCaseData> MetafileRenderingData
        {
            get
            {
                yield return new TestCaseData(new XamlFixedSaveOptions(), MetafileRenderingMode.Bitmap);
                yield return new TestCaseData(new XpsSaveOptions(), MetafileRenderingMode.Vector);
                yield return new TestCaseData(new SwfSaveOptions(), MetafileRenderingMode.VectorWithFallback);
            }
        }

        public static IEnumerable<TestCaseData> NumeralFormatData
        {
            get
            {
                yield return new TestCaseData(new HtmlFixedSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new ImageSaveOptions(SaveFormat.Docx), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new PdfSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new PsSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new SvgSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new XamlFixedSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new XpsSaveOptions(), NumeralFormat.EasternArabicIndic);
                yield return new TestCaseData(new SwfSaveOptions(), NumeralFormat.European);
            }
        }

        public static IEnumerable<TestCaseData> PageCountData
        {
            get
            {
                yield return new TestCaseData(new HtmlFixedSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new ImageSaveOptions(SaveFormat.Docx), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new PdfSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new PsSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new SvgSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new XamlFixedSaveOptions(), NumeralFormat.ArabicIndic);
                yield return new TestCaseData(new XpsSaveOptions(), NumeralFormat.EasternArabicIndic);
                yield return new TestCaseData(new SwfSaveOptions(), NumeralFormat.European);
            }
        }

        [Test]
        [TestCaseSource(nameof(JpegQualityData))]
        public void JpegQualityDefaultValue(FixedPageSaveOptions objectSaveOptions)
        {
            Document doc = new Document();

            FixedPageSaveOptions saveOptions = objectSaveOptions;
            Assert.AreEqual(95, saveOptions.JpegQuality);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);
        }

        [Test]
        [TestCaseSource(nameof(MetafileRenderingData))]
        public void MetafileRendering(FixedPageSaveOptions objectSaveOptions, MetafileRenderingMode metafileRendering)
        {
            Document doc = new Document();

            FixedPageSaveOptions saveOptions = objectSaveOptions;
            saveOptions.MetafileRenderingOptions.RenderingMode = metafileRendering;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);
        }

        [Test]
        [TestCaseSource(nameof(NumeralFormatData))]
        public void RenderingOfNumerals(FixedPageSaveOptions objectSaveOptions, NumeralFormat numeralFormat)
        {
            Document doc = new Document();

            FixedPageSaveOptions saveOptions = objectSaveOptions;
            saveOptions.NumeralFormat = numeralFormat;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);
        }

        [Test]
        [TestCaseSource(nameof(PageCountData))]
        public void Num(FixedPageSaveOptions objectSaveOptions, int pageCount)
        {
            Document doc = new Document();

            FixedPageSaveOptions saveOptions = objectSaveOptions;
            saveOptions.PageCount = pageCount;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);
        }
    }
}
