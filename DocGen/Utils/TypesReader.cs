using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Utils
{
    class TypesReader
    {
        private List<string[]> types;
        public List<string[]> LoadTypes()
        {
            types = new List<string[]>();
            string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string assemblyFolder = System.IO.Path.GetDirectoryName(assemblyLocation);
            string path = System.IO.Path.Combine(assemblyFolder, "types.txt");
            //string path = @"types.txt";
            string[] lines;
            try
            {
                lines = System.IO.File.ReadAllLines(path);
                foreach (string line in lines)
                {
                    AddType(line);
                }
            }
            catch (Exception e)
            {
                // if file isn't read, ignore
            }

            return types;
        }

        private void AddType(string line)
        {
            string[] typeLine = line.Split(new Char[] { '\t' }, StringSplitOptions.None);
            types.Add(typeLine);

        }

        public string[] FindType(string type)
        {

            if (types == null || types.Count == 0)
            {
                return null;
            }

            foreach (string[] s in types)
            {
                if (type.Equals(s[0]))
                {
                    return s;
                }
            }

            return null;
        }

    }


}
