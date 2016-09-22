// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Text.RegularExpressions;

using Aspose.Words;
using Aspose.Words.Replacing;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRange : ApiExampleBase
    {
        #region Replace 

        [Test]
        public void ReplaceSimple()
        {
            //ExStart
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExSummary:Simple find and replace operation.
            // Open the document.
            Document doc = new Document(MyDir + "Range.ReplaceSimple.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello _CustomerName_,");

            // Check the document contains what we are about to test.
            Console.WriteLine(doc.FirstSection.Body.Paragraphs[0].GetText());

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = false;

            // Replace the text in the document.
            doc.Range.Replace("_CustomerName_", "James Bond", options); //instead of obsolete method doc.Range.Replace("_CustomerName_", "James Bond", false, false);

            // Save the modified document.
            doc.Save(MyDir + @"\Artifacts\Range.ReplaceSimple.doc");
            //ExEnd

            Assert.AreEqual("Hello James Bond,\r\x000c", doc.GetText());
        }

        [Test]
        public void ReplaceWithString()
        {
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This one is sad.");
            builder.Writeln("That one is mad.");

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = true;

            doc.Range.Replace("sad", "bad", options);
            
            doc.Save(MyDir + @"\Artifacts\ReplaceWithString.doc");
        }

        [Test]
        public void ReplaceWithRegex()
        {
            //ExStart
            //ExFor:Range.Replace(Regex, String)
            //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
            Document doc = new Document(MyDir + "Document.doc");
            doc.Range.Replace(new Regex("[s|m]ad"), "bad"); //this method is obsolete, but still continues to work
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\ReplaceWithRegex.doc");
        }

        [Test]
        public void ChangeTextToHyperlinks()
        {
            //ExStart
            //ExFor:Range
            //ExFor:ReplacingArgs.Match
            //ExFor:Range.Replace(Regex, string, FindReplaceOptions)
            //ExFor:ReplacingArgs.Replacement
            //ExFor:IReplacingCallback
            //ExFor:IReplacingCallback.Replacing
            //ExFor:ReplacingArgs
            //ExSummary: Shows how to replace text to hyperlinks
            Document doc = new Document(MyDir + @"Range.ChangeTextToHyperlinks.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("http://www.aspose.com");
            builder.Writeln("http://www.aspose.com/documentation/file-format-components/aspose.words-for-.net-and-java/index.html");

            // Create regular expression for URL search
            Regex regexUrl = new Regex(@"(?<Protocol>\w+):\/\/(?<Domain>[\w.]+\/?)\S*(?x)");

            FindReplaceOptions options = new FindReplaceOptions();
            options.Direction = FindReplaceDirection.Backward;
            options.ReplacingCallback = new ChangeTextToHyperlinksEvaluator(doc);

            // Run replacement, using regular expression and evaluator.
            doc.Range.Replace(regexUrl, "", options); //instead of obsolete method doc.Range.Replace(System.Text.RegularExpressions.Regex, Aspose.Words.Replacing.IReplacingCallback, bool)

            // Save updated document.
            doc.Save(MyDir + @"\Artifacts\Range.ChangeTextToHyperlinks.docx");
        }

        private class ChangeTextToHyperlinksEvaluator : IReplacingCallback
        {
            internal ChangeTextToHyperlinksEvaluator(Document doc)
            {
                this.mBuilder = new DocumentBuilder(doc);
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is the run node that contains the found text. Note that the run might contain other 
                // text apart from the URL. All the complexity below is just to handle that. I don't think there
                // is a simpler way at the moment.
                Run run = (Run)e.MatchNode;

                Paragraph para = run.ParentParagraph;

                string url = e.Match.Value;

                FindReplaceOptions options = new FindReplaceOptions();
                options.MatchCase = true;
                options.FindWholeWordsOnly = true;

                // We are using \xbf (inverted question mark) symbol for temporary purposes.
                // Any symbol will do that is non-special and is guaranteed not to be presented in the document.
                // The purpose is to split the matched run into two and insert a hyperlink field between them.
                para.Range.Replace(url, "\xbf", options);

                Run subRun = (Run)run.Clone(false);
                int pos = run.Text.IndexOf("\xbf");
                subRun.Text = subRun.Text.Substring(0, pos);
                run.Text = run.Text.Substring(pos + 1, run.Text.Length - pos - 1);

                para.ChildNodes.Insert(para.ChildNodes.IndexOf(run), subRun);

                this.mBuilder.MoveTo(run);

                // Specify font formatting for the hyperlink.
                this.mBuilder.Font.Color = Color.Blue;
                this.mBuilder.Font.Underline = Underline.Single;

                // Insert the hyperlink.
                this.mBuilder.InsertHyperlink(url, url, false);

                // Clear hyperlink formatting.
                this.mBuilder.Font.ClearFormatting();

                // Let's remove run if it is empty.
                if (run.Text.Equals(""))
                    run.Remove();

                // No replace action is necessary - we have already done what we intended to do.
                return ReplaceAction.Skip;
            }

            private readonly DocumentBuilder mBuilder;
        }
        //ExEnd

        #endregion

        [Test]
        public void ReplaceNumbersAsHex()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Write("There are few numbers that should be converted to HEX and highlighted: 123, 456, 789 and 17379.");

            FindReplaceOptions options = new FindReplaceOptions();
            
            // Highlight newly inserted content.
            options.ApplyFont.HighlightColor = Color.DarkOrange;
            options.ReplacingCallback = new NumberHexer();

            int count = doc.Range.Replace(new Regex("[0-9]+"), "", options);
        }

        // Customer defined callback.
        private class NumberHexer : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                // Parse numbers.
                int number = Convert.ToInt32(args.Match.Value);
                
                // And write it as HEX.
                args.Replacement = string.Format("0x{0:X}", number);

                return ReplaceAction.Replace;
            }
        }

        [Test]
        public void DeleteSelection()
        {
            //ExStart
            //ExFor:Node.Range
            //ExFor:Range.Delete
            //ExSummary:Shows how to delete a section from a Word document.
            // Open Word document.
            Document doc = new Document(MyDir + "Range.DeleteSection.doc");

            // The document contains two sections. Each section has a paragraph of text.
            Console.WriteLine(doc.GetText());

            // Delete the first section from the document.
            doc.Sections[0].Range.Delete();

            // Check the first section was deleted by looking at the text of the whole document again.
            Console.WriteLine(doc.GetText());
            //ExEnd

            Assert.AreEqual("Hello2\x000c", doc.GetText());
        }
        
        [Test]
        public void RangesGetText()
        {
            //ExStart
            //ExFor:Range
            //ExFor:Range.Text
            //ExId:RangesGetText
            //ExSummary:Shows how to get plain, unformatted text of a range.
            Document doc = new Document(MyDir + "Document.doc");
            string text = doc.Range.Text;
            //ExEnd
        }

        [Test]
        public void RangesDeleteText()
        {
            //ExStart
            //ExId:RangesDeleteText
            //ExSummary:Shows how to delete all characters of a range.
            Document doc = new Document(MyDir + "Document.doc");
            doc.Sections[0].Range.Delete();
            //ExEnd
        }
    }
}
