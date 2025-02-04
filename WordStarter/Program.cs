using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace WordStarter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Word.Application word = new Word.Application();
            word.Visible = true;
            var executingAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var doc1 = Path.Combine(executingAssemblyDirectory, "..", "..", "..", "DocumentLevelAddin", "bin", "debug", "SampleDocument.docx");
            Console.WriteLine("Opening document ...");
            var wordDocument1 = word.Documents.Open(doc1);

            Console.WriteLine("Press ENTER to continue ...");
            Console.ReadKey();

            wordDocument1.Close();
            word.Quit();
        }
    }
}
