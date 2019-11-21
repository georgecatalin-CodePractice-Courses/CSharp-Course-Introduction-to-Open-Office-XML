using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;


namespace ConsoleApp_TryOOXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathToFolder = @"C:\Demo";
            string fileName = @"\Presentation.pptx";
            string completePathToFile = pathToFolder + fileName;

            Console.WriteLine(PPTGetSlideCount(completePathToFile));

            Console.WriteLine(PPTGetSlideCount(completePathToFile, false));
            Console.ReadLine();
        }


        //using System;
        //using System.Linq;

        // Return the number of slides, including hidden slides.
        public static int PPTGetSlideCount(string fileName)
        {
            return PPTGetSlideCount(fileName, true);
        }

        public static int PPTGetSlideCount(string fileName, bool includeHidden)
        {
            int slidesCount = 0;

            using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
            {
                // Get the presentation part of the document.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart != null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.SlideParts.Count();
                        //slidesCount = presentationPart.GetPartsCountOfType<SlidePart>(); //-->used for SDK 2.0
                    }
                    else
                    {
                        // Each slide can include a Show property, which if hidden will contain the value "0".
                        // The Show property may not exist, and most likely will not, for non-hidden slides.

                        var slides = presentationPart.SlideParts.
                         Where((s) => (s.Slide != null) &&
                           ((s.Slide.Show == null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        slidesCount = slides.Count();

                        /* *** Following code used for SDK 2.0 *** */
                        //var slides = presentationPart.GetPartsOfType<SlidePart>().
                        //  Where((s) => (s.Slide != null) &&
                        //    ((s.Slide.Show == null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        //slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }

    }
}
