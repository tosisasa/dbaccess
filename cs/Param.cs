using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;

namespace XXX
{
    class Param
    {
        static Dictionary<string, string> dict = null;

        public static String getValue(String key)
        {
            if (dict == null)
            {
                readFile();
            }

            if (!dict.ContainsKey(key))
            {
                return "";
            }
            return dict[key];
        }

        static void readFile(){
            dict = new Dictionary<string, string>();

            if (File.Exists("param.txt"))
            {
                string line = "";

                using (StreamReader sr = new StreamReader(
                    "param.txt", Encoding.GetEncoding("Shift_JIS")))
                {

                    while ((line = sr.ReadLine()) != null)
                    {
                        String[] list = line.Split('=');
                        dict.Add(list[0], list[1]);
                    }
                }
            }
        }
    }
}
