using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
//using System.IO;

namespace Scanorama
{
    class SlidesManipulation
    {
        public static void createPresentation(List<KeyValuePair<string, float>> titles)
        {
            Microsoft.Office.Interop.PowerPoint.Application oPowerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
            oPowerPoint.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            System.Console.WriteLine("PowerPoint application created.");
            Presentations oPres = oPowerPoint.Presentations;
            Presentation oPre = oPres.Add(MsoTriState.msoTrue);
            Slides oSlides = oPre.Slides;

            int nr = 1;
            foreach (var title in titles)
            {
                //Console.WriteLine(title.Key+" - "+ title.Value);
                createNewSlide(oSlides, nr, title.Key, title.Value);
                nr += 1;
            }
            Console.WriteLine("Checking for three lines...");
            TwoLines.twoLinesFilter(oSlides);
            //additionally to avoid three lines
            Console.WriteLine("Additionally checking for three lines...");
            TwoLines.twoLinesFilter(oSlides);
            FilesController.saveSlides(oPre);
            //oPre.Close();
            //oPowerPoint.Quit();
            //System.Console.WriteLine("PowerPoint application quitted.");

            //Clean up the unmanaged COM resource.
            if (oSlides != null)
            {
                Marshal.FinalReleaseComObject(oSlides);
                oSlides = null;
            }
            if (oPre != null)
            {
                Marshal.FinalReleaseComObject(oPre);
                oPre = null;
            }
            if (oPres != null)
            {
                Marshal.FinalReleaseComObject(oPres);
                oPres = null;
            }
            if (oPowerPoint != null)
            {
                Marshal.FinalReleaseComObject(oPowerPoint);
                oPowerPoint = null;
            }
        }

        public static void createEmptySlide(Slides oSlides, int slideNumber)
        {
            Slide oSlide = oSlides.Add(slideNumber, PpSlideLayout.ppLayoutBlank);
            oSlide.FollowMasterBackground = MsoTriState.msoFalse;
            oSlide.Background.Fill.ForeColor.RGB = colorizing(System.Windows.Media.Colors.Black);
        }

        public static void createNewSlide(Slides oSlides, int slideNumber, string slideText, float advanceSec)
        {
            /*Slide oSlide=null;
            if (slideText == "")
            {
                oSlide = oSlides.Add(slideNumber, PpSlideLayout.ppLayoutBlank);
                oSlide.FollowMasterBackground = MsoTriState.msoFalse;
                oSlide.Background.Fill.ForeColor.RGB = colorizing(System.Windows.Media.Colors.Black);
            }
            else
            {*/
           // Slide oSlide = oSlides.Add(slideNumber, PpSlideLayout.ppLayoutText);
            Slide oSlide = oSlides.Add(slideNumber, PpSlideLayout.ppLayoutTitleOnly);

            oSlide.FollowMasterBackground = MsoTriState.msoFalse;
            oSlide.Background.Fill.ForeColor.RGB = colorizing(System.Windows.Media.Colors.Black);

            Microsoft.Office.Interop.PowerPoint.Shapes oShapes = oSlide.Shapes;
            Microsoft.Office.Interop.PowerPoint.Shape oShape = oShapes[1];
            Microsoft.Office.Interop.PowerPoint.TextFrame oTxtFrame = oShape.TextFrame;
            
            TextRange oTxtRange = oTxtFrame.TextRange;
            slideText = setItalic(slideText, oTxtRange);
            oTxtRange.Text = slideText;

            oTxtRange.Font.Size = 44;
            oTxtRange.Font.Name = "Arial";
            oTxtRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            oTxtRange.Font.Color.RGB = colorizing(System.Windows.Media.Colors.White);

            //repositing text in the shape does not work
            //oTxtFrame.MarginTop = 10;
            //oShape.Top = 2;
            // }

            //only if advanced not after 0 sec, it turns on, AdvanceOnClick is still true
            if (advanceSec != 0)
            {
                oSlide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
                oSlide.SlideShowTransition.AdvanceTime = advanceSec;
                //oSlide.SlideShowTransition.AdvanceOnClick = MsoTriState.msoTrue;
            }
        }

        public static string setItalic(string slideText, TextRange textRange)
        {
            if (slideText.Contains("<i"))
            {
                Console.WriteLine("detected"+ textRange.Runs().Characters().Count);
                /*List<string> texts= new List<string>();
                string text = slideText;

                while (text.Contains("<i>")) {
                    int st = text.IndexOf("<i>") + 3;
                    int ls = text.IndexOf("</i>", st);
                    texts.Add(slideText.Substring(0, st));
                    texts.Add(slideText.Substring(st, ls));
                    text = text.Substring(ls + 4);
                }

                foreach (string s in texts) {
                    textRange.Runs()
                }*/
               MsoTriState state = MsoTriState.msoFalse;
               foreach (TextRange tr in textRange.Characters()) {
                    if (tr.Text == ">") {
                        state = changeState(state);
                        Console.WriteLine(state);
                        tr.Font.Italic = state;
                    }
                }
                slideText = slideText.Replace("<i>", "");
                slideText = slideText.Replace("</i>", "");
                //textRange.Font.Italic = MsoTriState.msoTrue;
            }
            if (slideText.Contains("#"))
            {
                /*int c = 0;
                for (int e = 0; e < s.Length; e++)
                {
                    if (slideText[e] == '#')
                    {
                        c++;
                    }
                }*/
                slideText = slideText.Replace("#", "");
                textRange.Font.Italic = MsoTriState.msoTrue;
            }
            return slideText;
        }

        public static MsoTriState changeState(MsoTriState state)
        {
            if (state == MsoTriState.msoTrue)
            {
                return MsoTriState.msoFalse;
            }
            return MsoTriState.msoTrue;
        }

        public static int colorizing(System.Windows.Media.Color color)
        {
            int iColor = color.R + 0xFF * color.G + 0xFFFF * color.B;
            return iColor;
        }
    }   
}
