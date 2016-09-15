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
    internal class ExHtmlLoadOptions : ApiExampleBase
    {
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void SupportVml(bool supportVml)
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportVml = supportVml;
            loadOptions.WebRequestTimeout = 1000;

            Document doc = new Document(MyDir + "", loadOptions);
        }

        public void WebRequestTimeoutDefaultValue()
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            Assert.Equals(100000, loadOptions.WebRequestTimeout);
        }
    }
}
