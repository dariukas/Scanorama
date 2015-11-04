using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace Scanorama
{
    class Main
    {
        public static void Run()
        {
            Prepare();
            Environment.Exit(0);
        }

        public static void Prepare()
        {
            SlidesManipulation.createPresentation(TitlesManipulation.readTitlesFromFile(FilesController.openFile()));
            //MainSpeech.Run();
        }
    }
}
