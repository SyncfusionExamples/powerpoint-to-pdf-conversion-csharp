using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint presentation file stream.
using (FileStream inputStream = new FileStream(@"../../../Data/Template.pptx", FileMode.Open, FileAccess.Read))
{
    //Open the existing PowerPoint presentation with loaded stream.
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
        //Add fallback font for specific unicode range.
        // Arabic.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0600, 0x06ff, "Arial"));
        // Hebrew.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0590, 0x05ff, "Times New Roman"));
        // Hindi.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0900, 0x097F, "Nirmala UI"));
        // Chinese.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x4E00, 0x9FFF, "DengXian"));
        // Japanese.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x3040, 0x309F, "MS Gothic"));
        // Thai.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0E00, 0x0E7F, "Tahoma"));
        // Korean.
        pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0xAC00, 0xD7A3, "Malgun Gothic"));
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