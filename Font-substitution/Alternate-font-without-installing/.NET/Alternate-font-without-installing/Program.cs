﻿using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint presentation file stream. 
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open,  FileAccess.ReadWrite))
{
    //Load an existing PowerPoint Presentation. 
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
        //Hook the font substitution event to handle unavailable fonts.
        //This event will be triggered when a font used in the document is not found in the production environment.
        pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
        //Convert PowerPoint into PDF document. 
        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
        {
           //Unhook the font substitution event after the conversion is complete.
            pptxDoc.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
            //Save the PDF file to file system. 
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
            {
                pdfDocument.Save(outputStream);
            }
        }
    }
}
/// <summary>
/// Sets the alternate font stream when a specified font is unavailable in the production environment.
/// </summary>
/// <param name="sender">FontSettings type of the Presentation in which the specified font stream is used but unavailable in production environment. </param>
/// <param name="args">Retrieves the unavailable font name and receives the substitute font stream for conversion. </param>
static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
{
    //Check if the original font is "Arial Unicode MS" and substitute with alternate font.
    if (args.OriginalFontName == "Arial Unicode MS" && args.FontStyle == FontStyle.Bold)
        args.AlternateFontStream = new FileStream(Path.GetFullPath(@"../../../Data/cambriab.ttf"), FileMode.Open, FileAccess.ReadWrite);
    else if (args.OriginalFontName == "Arial Unicode MS" && args.FontStyle == FontStyle.Regular)
        args.AlternateFontStream = new FileStream(Path.GetFullPath(@"../../../Data/BROADW.TTF"), FileMode.Open, FileAccess.ReadWrite);
    else
        args.AlternateFontStream = new FileStream(Path.GetFullPath(@"../../../Data/COOPBL.TTF"), FileMode.Open, FileAccess.ReadWrite);
}