using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Windows.Media;
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
            
            oTxtRange.Text = slideText;
           // oTxtRange = setItalic(oTxtRange);

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

        public static TextRange setItalic(TextRange textRange)
        {
            string text = textRange.Text;
            if (text.Contains("<i"))
            {
               MsoTriState state = MsoTriState.msoFalse;
               foreach (TextRange tr in textRange.Characters()) {
                    tr.Font.Italic = state;
                    if (tr.Text == ">")
                    {
                        state = changeState(state);
                    }
                }

                foreach (TextRange trr in textRange.Runs())
                {
                    Console.WriteLine("RUN: "+trr.Text);
                    trr.Text = trr.Text.Replace("<i>", "").Replace("</i>", "");
                }
                return textRange;
            }

            if (text.Contains("#"))
            {
                MsoTriState state = MsoTriState.msoFalse;
                foreach (TextRange tr in textRange.Characters())
                {
                    tr.Font.Italic = state;
                    if (tr.Text == "#")
                    {
                        state = changeState(state);
                        tr.Text = "";
                    }
                }
                return textRange;
            }
            /* int c = 0;
             for (int e = 0; e < s.Length; e++)
             {
                 if (slideText[e] == '#')
                 {
                     c++;
                 }
             }*/
            return null;
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

            if (color == System.Windows.Media.Colors.White) {
                return 16777215;

            }


            int iColor = color.R + 0xFF * color.G + 0xFFFF * color.B;
            string myHex = new ColorConverter().ConvertToString(iColor);
            //Console.WriteLine(iColor+" spalva R: "+color.R+" G: "+color.G+" B: "+color.B+" "+ myHex);
            //string.Format("#{0:X2}{1:X2}{2:X2}{3:X2}", color.A, color.R, color.G, color.B);
            //Color colorAfter = (Color)ColorConverter.ConvertFromString(myHex);
            //Color colorAfter = ConvertStringToColor(myHex);
            //Console.WriteLine(iColor + " spalva R: " + colorAfter.R + " G: " + colorAfter.G + " B: " + colorAfter.B + " " + myHex);
           iColor = Int32.Parse(myHex);





            return iColor;
        }

        public static System.Windows.Media.Color ConvertStringToColor(String hex)
        {
            //remove the # at the front
            hex = hex.Replace("#", "");

            byte a = 255;
            byte r = 255;
            byte g = 255;
            byte b = 255;

            int start = 0;

            //handle ARGB strings (8 characters long)
            if (hex.Length == 8)
            {
                a = byte.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                start = 2;
            }

            //convert RGB characters to bytes
            r = byte.Parse(hex.Substring(start, 2), System.Globalization.NumberStyles.HexNumber);
            g = byte.Parse(hex.Substring(start + 2, 2), System.Globalization.NumberStyles.HexNumber);
            b = byte.Parse(hex.Substring(start + 4, 2), System.Globalization.NumberStyles.HexNumber);

            return System.Windows.Media.Color.FromArgb(a, r, g, b);
        }


    }   
}
