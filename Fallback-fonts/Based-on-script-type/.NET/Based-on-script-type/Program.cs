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
        //Add fallback font for "Arabic" script type.
        pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Arabic, "Arial, Times New Roman");
        //Add fallback font for "Hebrew" script type.
        pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Hebrew, "Arial, Courier New");
        //Add fallback font for "Chinese" script type.
        pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Chinese, "DengXian, MingLiU");
        //Add fallback font for "Japanese" script type.
        pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Japanese, "Yu Mincho, MS Mincho");
        //Add fallback font for "Thai" script type.
        pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Thai, "Tahoma, Microsoft Sans Serif");
        //Add fallback font for "Korean" script type.
        pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic, Batang");
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