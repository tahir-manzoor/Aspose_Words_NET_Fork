﻿using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;

namespace _05._03_MovingtheCursor
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Shows how to access the current node in a document builder.
            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;

            // Shows how to move a cursor position to a specified node.
            builder.MoveTo(doc.FirstSection.Body.LastParagraph);

            // Shows how to move a cursor position to the beginning or end of a document.
            builder.MoveToDocumentEnd();
            builder.Writeln("This is the end of the document.");

            builder.MoveToDocumentStart();
            builder.Writeln("This is the beginning of the document.");

            doc.Save("outputDocument.doc");
        }
    }
}
