﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class DocumentTests : TestBase
    {
        [TestMethod]
        public void CreateUsingFile()
        {
            var filename = Path.Combine(TempDirectory, "CreateUsingFile.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            
            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}