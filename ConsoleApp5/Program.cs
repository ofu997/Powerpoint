using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// added these
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;


namespace Powerpoint
{
    /*
     Create a solution that accepts user input and generates a power point slide with; 
        Title area 
        Text area 
        and a image suggestion are that utilizes words in the title, and bold words in the text are to bring suggested images in, with ability to select multiple images to included in the slide 
        Have them make windows form to provide the solution. 
    */

    class Program
    {
        public static MsoTextOrientation MsoTextOrientationHorizontal { get; private set; }

        static void Main(string[] args)
        {
            string Title = "";
            string Text = "";
            string ImageURL = "";
            Console.WriteLine("Enter the name of the slide title, please.");
            Title += Console.ReadLine().ToString();
            Console.WriteLine("Enter image search key words, please.");
            Text += Console.ReadLine().ToString();
            
            Application pptApplication = new Application();
            Presentation pptpresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;
            
            // Microsoft.Office.Interop.PowerPoint.TextFrame textFrame;

            Microsoft.Office.Interop.PowerPoint.CustomLayout custLayout =
                pptpresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Creates new slide
            slides = pptpresentation.Slides;
            slide = slides.AddSlide(1, custLayout);

            Microsoft.Office.Interop.PowerPoint.TextRange objText;
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Font.Size = 55;
            objText.Text = Title;

            Microsoft.Office.Interop.PowerPoint.TextRange objText2;
            objText2= slide.Shapes[2].TextFrame.TextRange;
            ImageURL += "https://www.google.com/search?q=" + Text + "&tbm=isch";
            objText2.Text = ImageURL;
                objText2.Font.Name = "Book Angigua";
                objText2.Font.Size = 12;
                objText2.ActionSettings[Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = ImageURL;
            
            pptpresentation.SaveAs("newslide.pptx",
                Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                Microsoft.Office.Core.MsoTriState.msoTrue);
        }
    }
}