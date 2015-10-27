using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Scanorama
{
    class TitlesManipulation
    {
        public static SortedDictionary<string, float> readTitlesFromFile(string fileName)
        {
            System.Console.WriteLine("Reading the titles from the file {0}...", fileName);
            SortedDictionary<string, float> titles = new SortedDictionary<string, float>();
            //string title = System.IO.File.ReadAllText(fileName);
            string[] lines = System.IO.File.ReadAllLines(fileName);

            float duration = 0;
            float emptyDuration = 0;
            string text = "";//accumulating texts
            string timecode = "00:00:00,00";
            foreach (string line in lines)
            {
                if (line == "")
                {
                    titles.Add("", emptyDuration);
                    titles.Add(text, duration);
                    text = "";
                }
                else if (line.Contains(">"))
                {
                    string[] timecodes = line.Replace('|', '\u0020').Trim().Split('>');
                    duration =calculateDuration(timecodes[0], timecodes[1]);
                    emptyDuration = calculateDuration(timecode, timecodes[0]);
                    timecode = timecodes[1];
                }
                else if (Regex.IsMatch(line, @"^\D"))
                //else if(!char.IsDigit(line[0]))
                {
                    text = text + line.Trim() + "\n";
                }
          
                //Console.WriteLine("\t" + line);
                //Console.WriteLine("Press any key to exit.");
                //System.Console.ReadKey();
            }
            //System.Console.WriteLine("The content is {0}", titles[1]);
            return titles;
        }

        public static float calculateDuration(string start, string finish) {
               return Convert.ToSingle(timecodeToSeconds(finish) - timecodeToSeconds(start));
        }

        public static double timecodeToSeconds(string value)
        {
            //string value1 = "00:01:02,480";
            value.Replace(',', '.');
            TimeSpan span = TimeSpan.Parse(value);
            double seconds = span.TotalSeconds;
            return seconds;
        }

    }
}
