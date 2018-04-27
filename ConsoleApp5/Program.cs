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
    /*
     Create a solution that accepts user input and generates a power point slide with; 
        Title area 
        Text area 
        and a image suggestion are that utilizes words in the title, and bold words in the text are to bring suggested images in, with ability to select multiple images to included in the slide 
        Have them make windows form to provide the solution. 

        - it says something about bold words in the text to bring images in, I'm not sure exactly what you mean by this
       if the text has bold items, then only those will be used in the search criteria. 
- Should the list of pictures come from the local machine?  Or try to search google images maybe (never done anything like this so will probably take some time to figure it out)? 
        this is a list from the internet, easily done using WebBrowser control and google search. Get creative that is the point. 
    */

    class Program
    {
        //public static MsoTextOrientation MsoTextOrientationHorizontal { get; private set; }

        static void Main(string[] args)
        {
            /*
            string Title = "";
            string Text = "";
            //string ImageURL = "";
            Console.WriteLine("Enter the name of the slide title, please.");
            Title += Console.ReadLine().ToString();
            Console.WriteLine("Enter image search key words, please.");
            Text += Console.ReadLine().ToString();

            Application pptApplication = new Application();
            Presentation pptpresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;

            

            Microsoft.Office.Interop.PowerPoint.CustomLayout custLayout =
                pptpresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Creates new slide
            slides = pptpresentation.Slides;
            slide = slides.AddSlide(1, custLayout);

            Microsoft.Office.Interop.PowerPoint.TextRange objText;
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Font.Size = 55;
            objText.Text = Title;

            //
            Microsoft.Office.Interop.PowerPoint.TextRange objText2;
            objText2= slide.Shapes[2].TextFrame.TextRange;
            ImageURL += "https://www.google.com/search?q=" + Text + "&tbm=isch";
            objText2.Text = Text;
                objText2.Font.Name = "Book Angigua";
                objText2.Font.Size = 12;
                objText2.ActionSettings[Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = ImageURL;
            
            //


            //string boldWordsURL = "";
            string[] boldWords = Title.Split(' ');
            //for(int i=0; i<boldWords.Length;i++)
            //{
            //string[] textRangeBoldWords = Text.Split(' ');
            //Console.WriteLine(boldWords[i]);



            //
                Microsoft.Office.Interop.PowerPoint.TextRange objText1;
                boldWordsURL += "https://www.google.com/search?q=" + boldWords[0] + "&tbm=isch";
                objText1 = slide.Shapes[2].TextFrame.TextRange;
                objText1.Text = boldWords[0];
                objText1.Font.Name = "Book Angigua";
                objText1.Font.Size = 16;
                objText1.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                objText1.ActionSettings[Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = boldWordsURL;

                Microsoft.Office.Interop.PowerPoint.TextRange objText2;
                boldWordsURL += "https://www.google.com/search?q=" + boldWords[1] + "&tbm=isch";
                objText2 = slide.Shapes[2].TextFrame.TextRange;
                objText2.Text = boldWords[1];
                objText2.Font.Name = "Book Angigua";
                objText2.Font.Size = 16;
                objText2.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                objText2.ActionSettings[Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = boldWordsURL;
                //

            //}

            slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
               Left: 100,
               Top: 100,
              Width: 200,
              Height: 100
              ).TextFrame.TextRange.Text = Text;

            for (int i = 0; i < boldWords.Length; i++)
            {
                string ImageURL = "https://www.google.com/search?q=" + boldWords[i] + "&tbm=isch" + "\r";

                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                   Left: 100,
                   Top: 100*(i+2),
                  Width: 200,
                  Height: 100
                  ).TextFrame.TextRange.Text = ImageURL;
            }

            //
            if (boldWords[1] != null)

                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: 0,
                    Top: 0,
                   Width: 200,
                   Height: 100
                                
                   ).TextFrame.TextRange.Text = boldWords[1];
            
            TextFrame.TextRange.ActionSettings
           [Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick]
           .Hyperlink.Address = boldWords[1];

            
    slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
Left: 0,
Top: 0,
Width: 200,
Height: 0
)
.TextFrame.TextRange.Text = boldWords[1];
//


            pptpresentation.SaveAs("newslide.pptx",
                Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                Microsoft.Office.Core.MsoTriState.msoTrue);
            */

            string Title = "";
            string Text = "";
            string ImageURL = "";
            Console.WriteLine("Enter the name of the slide title, please.");
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
                
                /*
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
               Left: 300,
               Top: 50,
              Width: 100,
              Height: 200
              ).TextFrame.TextRange.Text = Title;
              */

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