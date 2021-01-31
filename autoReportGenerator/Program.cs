using System;
using Word = Microsoft.Office.Interop.Word;

namespace autoReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initializing the word application
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            //creating a word document
            Word.Document doc = wordApp.Documents.Add();
            doc.Activate();

            //writing text to document
            doc.Content.Text += "Hello World!";

            //save
            doc.SaveAs2("test.docx");

            //close and quit
            doc.Close();
            wordApp.Quit();
        }
    }
}
