using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scanorama
{
    class Zodynas
    {
        public static ILookup<string, float> createDict(){
            List<KeyValuePair<string, float>> list = new List<KeyValuePair<string, float>>();
            list.Add(new KeyValuePair<string, float>("", 10));
            var dict = list.ToLookup(kvp => kvp.Key, kvp => kvp.Value);
            return dict;
        }

        public static void printDict(ILookup<string, float> dict)
        {
            //for ILookup
            foreach (var kvp in dict)
            {
                foreach (var value in kvp)
                {
                    Console.WriteLine("For " + kvp.Key + " is " + value);
                }
            }
        }
    }
}
