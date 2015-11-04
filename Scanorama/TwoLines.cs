using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace Scanorama
{
    class TwoLines
    {
        public static void twoLinesFilter(Slides slides) {

            foreach (Slide slide in slides)
            {
                Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[1];
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        var textRange = shape.TextFrame.TextRange;
                        //check if there are more than two lines, and needs additional divisions
                        if (textRange.Lines().Count > 2)
                        {
                            if (textRange.Sentences().Count > 1)
                            {
                                divideToSentences(slides, slide, textRange);
                            } else {
                                divideToLines(slides, slide, textRange);
                            }
                        }
                    }
                }
            }
        }

        public static void divideToLines(Slides slides, Slide slide, TextRange textRange)
        {
            var slideNumber = slide.SlideIndex;
            float slideDuration = slide.SlideShowTransition.AdvanceTime;
            int divisionNumber = textRange.Lines().Count;
            float duration = durationAfterDivisions(slideDuration, divisionNumber / 2);
            string textFrmLines = "";
            foreach (TextRange line in textRange.Lines())
            {
                if (textFrmLines.Length > 0)
                {
                    textFrmLines += line.Text;
                    SlidesManipulation.createNewSlide(slides, ++slideNumber, textFrmLines.Trim(), duration);
                    SlidesManipulation.createNewSlide(slides, ++slideNumber, "", 0.01F);
                    textFrmLines = "";
                }
                else
                {
                    textFrmLines += line.Text;
                }
            }
            //add the rest of textFrmLines
            if (textFrmLines.Length > 0)
            {
                SlidesManipulation.createNewSlide(slides, ++slideNumber, textFrmLines, duration);
                SlidesManipulation.createNewSlide(slides, ++slideNumber, "", 0.1F);
            }
                //delete slides
                slide.Delete();
                slides[slideNumber].Delete();
            }

        public static void divideToSentences(Slides slides, Slide slide, TextRange textRange) {
            var slideNumber = slide.SlideIndex;
            float slideDuration = slide.SlideShowTransition.AdvanceTime;
            //the number represents to how many slides to divide
            int divisionNumber = textRange.Sentences().Count;
            float duration = durationAfterDivisions(slideDuration, divisionNumber);
            foreach (TextRange sentence in textRange.Sentences())
            {
                SlidesManipulation.createNewSlide(slides, ++slideNumber, sentence.Text.Trim(), duration);
                SlidesManipulation.createNewSlide(slides, ++slideNumber, "", 0.01F);
            }
            //delete slides
            slide.Delete();
            slides[slideNumber].Delete();
        }
        
        public static void divideText(TextRange textRange)
        {
            string textPart = "";
            int m = 0;
            foreach (TextRange line in textRange.Lines())
            {
                if (m == 2 && line.Text.Contains(","))
                {
                    //Char.IsPunctuation;
                    //Char.getUnicodeCategory+
                    //https://msdn.microsoft.com/en-us/library/system.globalization.unicodecategory(v=vs.110).aspx
                    char[] separator = { ',' };
                    string[] lineParts = line.Text.Split(separator, 2);

                    textPart += lineParts[0];
                    //add to slide
                    textPart = lineParts[1];
                }
                m++;
                textPart += line.Text;
            }
            //add to slide
        }

        public static float durationAfterDivisions(float slideDuration, int divisionNumber)
        {
            float emptyDuration = 0.1F;
            return slideDuration / divisionNumber - emptyDuration;
        }
    }
}
