using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocGen.Model;
using DocGen.Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.Model
{
    class GroupFiller
    {
        private List<Group> groups = new List<Group>();

        public List<Group> FillGroups(List<Components> bom)
        {
            Group group = new Group();
            TypesReader types = new TypesReader();
            types.LoadTypes();
            string lastDes = null;
            string currentDes = null;
            string[] typeDescribe = null;
            //размещение элементов по типам
            foreach (Components c in bom)
            {
                // получить designator без чисел (DD вместо DD2)
                currentDes = DesType(c.GetDesignators());

                // если появился новый десигнатор
                if (!currentDes.Equals(lastDes))
                {
                    // add current group
                    if (!group.IsEmpty())
                    {
                        groups.Add(group);
                    }

                    // create new group and fill fields
                    group = new Group();

                    typeDescribe = types.FindType(currentDes);

                    if (typeDescribe != null)
                    {
                        group.DesignatorType = typeDescribe[0];
                        group.TypeDescription = typeDescribe[1];
                        group.TypeDescriptions = typeDescribe[2];
                    }
                } 

                group.Add(c); // add component in the same group
                lastDes = currentDes;

            }

            //add last group
            if (!group.IsEmpty())
            {
                groups.Add(group);
            }
            return groups;
        }

        private string DesType(string des)
        {
            if (String.IsNullOrEmpty(des))
            {
                return "";
            }
            Regex regex = new Regex(@"[a-zA-Z]+");
            Match match = regex.Match(des);
            string mathced = match.Groups[0].Value;

            if (!String.IsNullOrEmpty(mathced))
            {
                return mathced;
            }
            else
            {
                return des;
            }
        }
    }
}
