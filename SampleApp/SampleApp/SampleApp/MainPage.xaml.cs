using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Essentials;
using Xamarin.Forms;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace SampleApp
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
            
        }

        private const string fileName = "xamarinlibrary.docx";
        
        protected override void OnAppearing()
        {
            var dir = FileSystem.AppDataDirectory;
            var filepath = $"{dir}/{fileName}";
            if (System.IO.File.Exists(filepath))
            {
                System.IO.File.Delete(filepath);
            }            
        }

        private void ReadResourceFile_Clicked(object sender, EventArgs e)
        {
            var resourceID = "SampleApp.Files.XamarinLibraryWord.docx";
            var assembly = Assembly.Load("SampleApp");

            using (Stream stream= assembly.GetManifestResourceStream(resourceID))
            {
                var doc = WordprocessingDocument.Open(stream, false);

                Body body = doc.MainDocumentPart.Document.Body;
                var para = body.ElementAt(0) as Paragraph;
                var run = para.ElementAt(0) as Run;
                var text = run.ElementAt(0) as Text;
                ResourceValLabel.Text = $"Resource First Value:{text.Text}";
            }
        }

        private void CreateWordDoc(string filepath)
        {
            var dir = FileSystem.AppDataDirectory;
            filepath = $"{dir}/{filepath}";

            using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text("XamarinLibrary string created from openxml!"));
            }
        }

        private void CreateDocxBtn_Clicked(object sender, EventArgs e)
        {
            CreateWordDoc(fileName);
            CreateStatusLabel.Text = "Status:Create Finished";
        }

        private void OpenDocxBtn_Clicked(object sender, EventArgs e)
        {
            var dir = FileSystem.AppDataDirectory;
            var filepath = $"{dir}/{fileName}";
            var doc = WordprocessingDocument.Open(filepath, true);
            Body body = doc.MainDocumentPart.Document.Body;
            var para = body.ElementAt(0) as Paragraph;
            var run = para.ElementAt(0) as Run;
            var text = run.ElementAt(0) as Text;
            FirstValLabel.Text = $"First value:{text.Text}"; 
        }
    }
}
