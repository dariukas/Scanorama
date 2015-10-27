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
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = path + "\\MyTitles.txt";

            //Browse Files
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".txt";
            dlg.Filter = "TXT Files (*.txt)|*.txt|SRT Files (*.srt)|*.srt|DOC Files (*.doc)|*.doc|RTF Files (*.rtf)|*.rtf";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                fileName = dlg.FileName;
                //textBox1.Text = filename;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }
            return fileName;
        }

        public static void saveSlides(Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string fileName = path + "\\Slides.pptx";

            //Browse Files
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PPTX Files (*.pptx)|*.pptx";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                fileName = dlg.FileName;
                //textBox1.Text = filename;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }

            presentation.SaveAs(fileName,
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
            Microsoft.Office.Core.MsoTriState.msoTriStateMixed);
            System.Console.WriteLine("PowerPoint application saved in {0}.", fileName);
        }

    }
}
