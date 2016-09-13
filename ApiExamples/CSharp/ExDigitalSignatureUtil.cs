// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    using System.Drawing;

    [TestFixture]
    public class ExDigitalSignatureUtil : ApiExampleBase
    {
        //ToDo: Change "Document.Signed.doc" to "Document.doc"
        [Test]
        public void RemoveAllSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove every signature from a document.
            //By stream:
            Stream docStreamIn = new FileStream(MyDir + "Document.Signed.doc", FileMode.Open);
            Stream docStreamOut = new FileStream(MyDir + "Document.NoSignatures.FromStream.doc", FileMode.Create);

            DigitalSignatureUtil.RemoveAllSignatures(docStreamIn, docStreamOut);

            docStreamIn.Close();
            docStreamOut.Close();

            //By string:
            Document doc = new Document(MyDir + "Document.Signed.doc");
            string outFileName = MyDir + "Document.NoSignatures.FromString.doc";

            DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);
            //ExEnd
        }

        //ToDo: Change "Document.Signed.doc" to "Document.doc"
        [Test]
        public void LoadSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExSummary:Shows how to load signatures from a document by stream and by string.
            Stream docStream = new FileStream(MyDir + "Document.Signed.doc", FileMode.Open);

            // By stream:
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(docStream);
            docStream.Close();

            // By string:
            digitalSignatures = DigitalSignatureUtil.LoadSignatures(MyDir + "Document.Signed.doc");
            //ExEnd
        }

        [Test]
        public void SignEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, String, DateTime)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, String, DateTime)
            //ExSummary:Shows how to sign documents.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By String
            Document doc = new Document(MyDir + "Document.doc");
            string outputDocFileName = MyDir + @"\Artifacts\Document.Signed.doc";

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now);
            
            //By Stream
            Stream docInStream = new FileStream(MyDir + "Document.doc", FileMode.Open);
            Stream docOutStream = new FileStream(MyDir + @"\Artifacts\Document.Signed.doc", FileMode.OpenOrCreate);

            DigitalSignatureUtil.Sign(docInStream, docOutStream, ch, "My comment", DateTime.Now);
            //ExEnd

            docInStream.Dispose();
            docOutStream.Dispose();
        }

        [Test]
        [TestCase("Stream")]
        [TestCase("Document")]
        public void SingWithPasswordDecrypring(string source)
        {
            // Create certificate holder from a file.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            if (source == "Document")
            {
                //ByDocument
                Document doc = new Document(MyDir + "Document.doc");
                string outputDocFileName = MyDir + @"\Artifacts\Document.Signed.doc";

                // Digitally sign encrypted with "docPassword" document in the specified path.
                DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "Comment", DateTime.Now, "docPassword");

                // Open encrypted document from a file.
                Document signedDoc = new Document(outputDocFileName, new LoadOptions("docPassword"));

                // Check that encrypted document was successfully signed.
                DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
                if (signatures.IsValid && (signatures.Count > 0))
                {
                    Assert.Pass(); //The document was signed successfully
                }
            }
            else
            {
                //By Stream
                Stream docInStream = new FileStream(MyDir + "Document.doc", FileMode.Open);
                Stream docOutStream = new FileStream(MyDir + @"\Artifacts\Document.Signed.doc", FileMode.OpenOrCreate);

                // Digitally sign encrypted with "docPassword" document in the specified path.
                DigitalSignatureUtil.Sign(docInStream, docOutStream, ch, "Comment", DateTime.Now, "docPassword");

                // Open encrypted document from a file.
                Document signedDoc = new Document(docOutStream, new LoadOptions("docPassword"));

                // Check that encrypted document was successfully signed.
                DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
                if (signatures.IsValid && (signatures.Count > 0))
                {
                    Assert.Pass(); //The document was signed successfully
                }

                docInStream.Dispose();
                docOutStream.Dispose();
            }
        }

        [Test]
        [ExpectedException(typeof(ArgumentException))]
        public void NoArgumentsForSing()
        {
            DigitalSignatureUtil.Sign(String.Empty, String.Empty, null, String.Empty, DateTime.Now, String.Empty);
        }

        [Test]
        [ExpectedException(typeof(NullReferenceException))]
        public void NoCertificateForSign()
        {
            //ByDocument
            Document doc = new Document(MyDir + "Document.doc");
            string outputDocFileName = MyDir + @"\Artifacts\Document.Signed.doc";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, null, "Comment", DateTime.Now, "docPassword");
        }
        
    }
}