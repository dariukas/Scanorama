using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace Scanorama
{
    class NotUsed
    {
        public static void TestPresentation()
        {
            Application PowerPoint_App = new Application();
            Presentations multi_presentations = PowerPoint_App.Presentations;
            Presentation presentation = multi_presentations.Open(FilesController.openFile());
            string presentation_text = "";
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                foreach (var item in presentation.Slides[i + 1].Shapes)
                {
                    var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            //Console.WriteLine("sentence: " + textRange.Sentences().Length);

                            //check if there are more than two lines, and needs additional divisions
                            if (textRange.Lines().Count > 2)
                            {
                                int slideNumber = shape.Parent.SlideIndex;
                                float slideDuration = presentation.Slides[i + 1].SlideShowTransition.AdvanceTime;
                                presentation.Slides[i + 1].Delete();
                                //CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
                                if (textRange.Sentences().Count > 1)
                                {
                                    //the number represents to how many slides to divide
                                    int divisionNumber = textRange.Sentences().Count;
                                    float duration = TwoLines.durationAfterDivisions(slideDuration, divisionNumber);

                                    foreach (TextRange sentence in textRange.Sentences())
                                    {
                                        SlidesManipulation.createNewSlide(presentation.Slides, slideNumber++, sentence.Text.Trim(), duration);
                                    }
                                }
                                else
                                {
                                    int divisionNumber = textRange.Lines().Count;
                                    float duration = TwoLines.durationAfterDivisions(slideDuration, divisionNumber / 2);
                                    string textFrmLines = "";
                                    foreach (TextRange line in textRange.Lines())
                                    {
                                        if (textFrmLines.Length > 0)
                                        {
                                            textFrmLines += line.Text;
                                            SlidesManipulation.createNewSlide(presentation.Slides, slideNumber++, textFrmLines, duration);
                                            textFrmLines = "";
                                        }
                                        else
                                        {
                                            textFrmLines += line.Text;
                                        }
                                    }
                                }
                            }
                            var text = textRange.Text;
                            presentation_text += text + " ";
                        }
                    }
                }
            }
            PowerPoint_App.Quit();
            Console.WriteLine(presentation_text);
            FilesController.saveSlides(presentation);
        }
    }
}
