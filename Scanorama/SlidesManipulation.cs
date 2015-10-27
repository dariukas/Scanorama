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
        public static void createPresentation(SortedDictionary<string, float> titles) {
            Microsoft.Office.Interop.PowerPoint.Application oPowerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
            oPowerPoint.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            System.Console.WriteLine("PowerPoint application created.");
            Presentations oPres = oPowerPoint.Presentations;
            Presentation oPre = oPres.Add(MsoTriState.msoTrue);
            Slides oSlides = oPre.Slides;

            int nr = 1;
            foreach (string key in titles.Keys)
            {
                createNewSlide(oSlides, nr, key, titles[key]);
                nr += 1;
            }

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
            Slide oSlide = oSlides.Add(slideNumber, PpSlideLayout.ppLayoutText);

            oSlide.FollowMasterBackground = MsoTriState.msoFalse;
            oSlide.Background.Fill.ForeColor.RGB = colorizing(System.Windows.Media.Colors.Black);


            Microsoft.Office.Interop.PowerPoint.Shapes oShapes = oSlide.Shapes;
            Microsoft.Office.Interop.PowerPoint.Shape oShape = oShapes[1];
            Microsoft.Office.Interop.PowerPoint.TextFrame oTxtFrame = oShape.TextFrame;
            TextRange oTxtRange = oTxtFrame.TextRange;

            if (slideText.Contains("<i"))
            {
                int st = slideText.IndexOf(">") + 1;
                int ls = slideText.LastIndexOf("<");
                slideText = slideText.Substring(st, ls - st);
                oTxtRange.Font.Italic = MsoTriState.msoTrue;
            }

            oTxtRange.Text = slideText;

            oTxtRange.Font.Size = 44;
            oTxtRange.Font.Name = "Arial";
            oTxtRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            oTxtRange.Font.Color.RGB = colorizing(System.Windows.Media.Colors.White);
            // }

            //only if advanced not after 0 sec, it turns on, AdvanceOnClick is still true
            if (advanceSec != 0)
            {
                oSlide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
                oSlide.SlideShowTransition.AdvanceTime = advanceSec;
            }

        }

        public static int colorizing(System.Windows.Media.Color color)
        {
            int iColor = color.R + 0xFF * color.G + 0xFFFF * color.B;
            return iColor;
        }

    }
}
