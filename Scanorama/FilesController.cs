using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scanorama
{
    class FilesController
    {

        //open the titles file using the dialog
        public static string openFile()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = desktopPath + "\\MyTitles.txt";

            //Browse Files
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = desktopPath + "\\Scano";
            dlg.DefaultExt = ".txt";
            dlg.Filter = "TXT Files (*.txt)|*.txt|SRT Files (*.srt)|*.srt|DOC Files (*.doc)|*.doc|RTF Files (*.rtf)|*.rtf";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                filePath = dlg.FileName;
                //textBox1.Text = filename;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }
            return filePath;
        }

        public static void saveSlides(Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string filePath = desktopPath + "\\Slides.pptx";


            // Microsoft.Office.Interop.PowerPoint.FileConverter fc = new Microsoft.Office.Interop.PowerPoint.FileConverter();
            // if (fc.CanSave) { }
            //https://msdn.microsoft.com/en-us/library/system.windows.forms.savefiledialog(v=vs.110).aspx

            //Browse Files
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.InitialDirectory = desktopPath+"\\Scano";
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PPTX Files (*.pptx)|*.pptx";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                filePath = dlg.FileName;
                //textBox1.Text = filename;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }

            presentation.SaveAs(filePath,
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
            Microsoft.Office.Core.MsoTriState.msoTriStateMixed);
            System.Console.WriteLine("PowerPoint application saved in {0}.", filePath);
        }

    }
}
