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

namespace ApiExamples
{
    using System.Collections.Generic;
    using System.IO;

    [TestFixture]
    internal class ExFixedPageSaveOptions : ApiExampleBase
    {
        public IEnumerable<TestCaseData> CountEqualsZeroAndHouseGrossIsGreaterTestCases
        {
            get
            {
                yield return new TestCaseData(new XamlFixedSaveOptions());
                yield return new TestCaseData(new XpsSaveOptions());
            }
        }
        
        [Test]
        [TestCaseSource("CountEqualsZeroAndHouseGrossIsGreaterTestCases")]
        public void JpegQualityDefaultValue(FixedPageSaveOptions objectSaveOptions)
        {
            FixedPageSaveOptions saveOptions = objectSaveOptions;
            Assert.AreEqual(95, saveOptions.JpegQuality);
        }
    }
}
