using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint presentation file stream.
using (FileStream fileStreamInput = new FileStream(@"../../../Data/Template.pptx", FileMode.Open, FileAccess.Read))
{
    //Open the existing PowerPoint presentation with loaded stream.
    using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
    {
        //Initialize the default fallback fonts collection.
        pptxDoc.FontSettings.FallbackFonts.InitializeDefault();
        //Convert PowerPoint into PDF document. 
        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
        {
            //Save the PDF file to file system. 
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
            {
                pdfDocument.Save(outputStream);
            }
        }
    }
}