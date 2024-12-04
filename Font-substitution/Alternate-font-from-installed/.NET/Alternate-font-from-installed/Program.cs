using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint presentation file stream. 
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing PowerPoint Presentation. 
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
        //Hook the font substitution event to handle unavailable fonts.
        //This event will be triggered when a font used in the document is not found in the production environment.
        pptxDoc.FontSettings.SubstituteFont += SubstituteFont;
        //Convert PowerPoint into PDF document. 
        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
        {
            //Unhook the font substitution event after the conversion is complete.
            pptxDoc.FontSettings.SubstituteFont -= SubstituteFont;
            //Save the PDF file to file system. 
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
            {
                pdfDocument.Save(outputStream);
            }
        }
    }
}

/// <summary>
/// Sets the alternate font when a specified font is unavailable in the production environment.
/// </summary>
/// <param name="sender">FontSettings type of the Presentation in which the specified font is used but unavailable in production environment. </param>
/// <param name="args">Retrieves the unavailable font name and receives the substitute font name for conversion. </param>
static void SubstituteFont(object sender, SubstituteFontEventArgs args)
{
    //Check if the original font is "Arial Unicode MS" and substitute with "Arial".
    if (args.OriginalFontName == "Arial Unicode MS")
        args.AlternateFontName = "Arial";
    else
        //Subsitutue "Times New Roman" for any other missing fonts.
        args.AlternateFontName = "Times New Roman";
}