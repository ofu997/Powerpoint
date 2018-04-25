using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// added this one
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
            string Title;
            string Text;
            string ImageURL;
            Console.WriteLine("Enter the name of the slide title, please.");
            Title=Console.ReadLine().ToString();
            Console.WriteLine("Enter slide description, please.");
            Text = Console.ReadLine().ToString();
            ImageURL = "https://www.google.com/search?q=" + Text + "&tbm=isch";

            // string[] PictureFile = { @"C:\powerpoint\img.1.jpg", @"C:\powerpoint\img.1.jpg", @"C:\powerpoint\img.1.jpg" };

            Application pptApplication = new Application();

            // Creates presentation 
            Presentation pptpresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            //for (int i = 0; i < 1; i++)
            //{
                Microsoft.Office.Interop.PowerPoint.Slides slides;
                Microsoft.Office.Interop.PowerPoint._Slide slide;
                Microsoft.Office.Interop.PowerPoint.TextRange objText;
            // Microsoft.Office.Interop.PowerPoint.TextFrame textFrame;

                Microsoft.Office.Interop.PowerPoint.CustomLayout custLayout=
                    pptpresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
                
                // Creates new slide
                slides = pptpresentation.Slides;
                slide = slides.AddSlide(1, custLayout);

                objText = slide.Shapes[1].TextFrame.TextRange;
                objText.Text = Text ;
                objText.Font.Name = "Book Angigua";
                objText.Font.Size = 32;
            // TRY TO INSERT TEXT BOX IN SLIDE 
            
            slide.Shapes.AddTextbox(Orientation: MsoTextOrientationHorizontal,
                Left: 1,
                Top: 1,
               Width: 1,
               Height: 1
                ).TextFrame.TextRange.Text = ImageURL; 
            

                //Microsoft.Office.Interop.PowerPoint.Shape shape =
                   // slide.Shapes[2];
                // slide.Shapes.AddPicture(PictureFile[i], Microsoft.Office.Core.MsoTriState.msoFalse,
                    // Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);
            // }
            pptpresentation.SaveAs("newslide.pptx", 
                Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, 
                Microsoft.Office.Core.MsoTriState.msoTrue);
        }
    }
}
