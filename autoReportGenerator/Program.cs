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
            doc.Content.Text += "First bit of text\n";


            //read from a text file
            string text = System.IO.File.ReadAllText(@"test.txt");
            System.Console.WriteLine("Contents of WriteText.txt = {0}", text);

            Word.Paragraph paragraph1 = doc.Content.Paragraphs.Add();
            String heading1 = "Heading 1";
            paragraph1.Range.Text = "4.3.1";
            paragraph1.Range.set_Style(heading1);
            paragraph1.Range.InsertParagraphAfter();

            Word.Paragraph p2 = doc.Content.Paragraphs.Add();
            p2.Range.Text = text;


            //save
            doc.SaveAs2("test.docx");

            //close and quit
            doc.Close();
            wordApp.Quit();
        }
    }
}