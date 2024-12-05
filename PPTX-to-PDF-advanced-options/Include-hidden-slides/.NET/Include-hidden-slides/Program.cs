using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint presentation file stream. 
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing PowerPoint Presentation. 
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
        //Create an instance of PresentationToPdfConverterSettings.
        PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
        //Enables the inclusion of hidden slides during the conversion process.
        pdfConverterSettings.ShowHiddenSlides = true;
        //Convert PowerPoint into PDF document. 
        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings))
        {
            //Save the PDF file to file system. 
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
            {
                pdfDocument.Save(outputStream);
            }
        }
    }
}
