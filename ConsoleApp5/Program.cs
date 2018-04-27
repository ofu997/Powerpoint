using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// added these
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace Powerpoint
{
    class Program
    {
        static void Main(string[] args)
        {
            string Title = "";
            string Text = "";
            string ImageURL = "";
            Console.WriteLine("Enter the name of the slide title, please. This will allow you to search for related images.");
            Title += Console.ReadLine().ToString();
            Console.WriteLine("Enter description, please.");
            Text += Console.ReadLine().ToString();

            Application pptApplication = new Application();
            Presentation pptpresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            string[] titleArray = Title.Split(' ');

            for (int i = 0; i < titleArray.Length; i++)
            {
                Microsoft.Office.Interop.PowerPoint.CustomLayout custLayout =
                pptpresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
                Microsoft.Office.Interop.PowerPoint.Slides slides;
                Microsoft.Office.Interop.PowerPoint._Slide slide;

                slides = pptpresentation.Slides;
                slide = slides.AddSlide(i + 1, custLayout);

                ImageURL = "https://www.google.com/search?q=" + titleArray[i] + "&tbm=isch";
                
                Microsoft.Office.Interop.PowerPoint.TextRange objTextTitle;
                objTextTitle = slide.Shapes[1].TextFrame.TextRange;
                objTextTitle.Font.Size = 24;
                objTextTitle.Text = Title;

                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                  Left: 150,
                  Top: 200,
                  Width: 200,
                  Height: 300
              ).TextFrame.TextRange.Text = Text;

                Microsoft.Office.Interop.PowerPoint.TextRange objText;
                objText = slide.Shapes[2].TextFrame.TextRange;
                objText.Font.Size = 24;
                objText.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                objText.Text = "Images for " + titleArray[i];
                objText.ActionSettings
                    [Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick]
                    .Hyperlink.Address = ImageURL;
                objText.Text.PadLeft(100);
            }
            pptpresentation.SaveAs("newslide.pptx",
    Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
    Microsoft.Office.Core.MsoTriState.msoTrue);
        }
    }    
}