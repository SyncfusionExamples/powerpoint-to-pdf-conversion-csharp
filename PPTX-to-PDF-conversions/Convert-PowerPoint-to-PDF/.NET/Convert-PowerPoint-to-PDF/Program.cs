using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint presentation file stream. 
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing PowerPoint Presentation. 
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
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