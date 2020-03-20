using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using DocGen.Model;
using DocGen.Model.Documents.Comparers;
using System.Diagnostics;
using System.Text.RegularExpressions;
using DocGen.Model.Documents;

namespace DocGen.Model.Documents
{
    abstract class SWDocument : IDocument
    {
        private List<SWComponents> bom;
        //protected ComponentList componentList;
        protected List<string> notes;
        protected Settings settings;
        public bool isGenerated = false;

        protected int maxNoteLength = 14;
        protected int limitNoteRow = 60;

        SWBOMReader bomReader;
		
		private List<List<SWComponents>> groups = new List<List<SWComponents>>();
		private List<SWComponents> assemblies;
		private List<SWComponents> parts;
		private List<SWComponents> standartParts;
		private List<SWComponents> otherParts;
		private List<SWComponents> matetials;

        public SWDocument ()
        {
            SettingsFactory factory = new SettingsFactory();
            this.settings = factory.GetSettings();
        }

        public void Generate()
        {
            ReadBOM();
            if (bom != null)
            {
                FillGroups();
                SortGroups();
                CreateNotes();
                FillDocumentRows(assemblies, "Сборочные единицы");
				FillDocumentRows(parts, "Детали");
				FillDocumentRows(standartParts, "Стандартные изделия");
				FillDocumentRows(otherParts, "Прочие изделия");
				FillDocumentRows(matetials, "Материалы");
                GenerateNotes();
                //CombineRows();
                isGenerated = true;
            }
        }

        private void ReadBOM()
        {
            bomReader = new SWBOMReader();
            bom = bomReader.ReadBOM();
        }

        protected virtual void FillGroups()
        {
            //assemblies = bom.ToList();
            assemblies = bom.Where(c => c.DocumentSection.Equals("Сборочные единицы")).ToList();
            parts = bom.Where(c => c.DocumentSection.Equals("Детали")).ToList();
			standartParts = bom.Where(c => c.DocumentSection.Equals("Стандартные изделия")).ToList();
            otherParts = bom.Where(c => c.DocumentSection.Equals("Прочие изделия")).ToList();
            matetials = bom.Where(c => c.DocumentSection.Equals("Материалы")).ToList();
			
			groups.Add(assemblies);
			groups.Add(parts);
			groups.Add(standartParts);
			groups.Add(otherParts);
			groups.Add(matetials);
        }

        protected virtual void SortGroups()
        {
			assemblies = assemblies.OrderBy(c => c.Designation).ToList();
			parts = parts.OrderBy(c => c.Designation).ToList();
			standartParts = standartParts.OrderBy(c => c.Name).ToList();
			otherParts = otherParts.OrderBy(c => c.Name).ToList();
			matetials = matetials.OrderBy(c => c.Name).ToList();
        }


        private void CreateNotes()
        {
            notes = new List<string>();
            foreach (List<SWComponents> group in groups)
            {
                foreach (SWComponents c in group)
                {
                    if (!String.IsNullOrEmpty(c.Replacement))
                    {
                        if (!notes.Contains(c.Replacement))
                        {
                            notes.Add(c.Replacement);
                        }
                        c.ReplacementNumber = notes.IndexOf(c.Replacement) + 1;
                    }
                }
            }
        }

        protected virtual void GenerateNotes()
        {
        }

        protected virtual void FillDocumentRows(List<SWComponents> group, String section)
        {
        }

        public virtual List<IRow> CombineRows()
        {
            return new List<IRow>();
        }

        protected string[] SplitNote(string note)
        {
            string[] splitted = note.Split(' ');
            List<string> notes = new List<string>();
            string tmp = "";
            for (int i = 0; i < splitted.Length; i++)
            {
                if ((tmp.Length + splitted[i].Length) < maxNoteLength - 1)
                {
                    if (tmp.Length > 0)
                    {
                        tmp += " ";
                    }
                    tmp += splitted[i];
                }
                else
                {
                    notes.Add(tmp);
                    tmp = splitted[i];
                }
            }
            if (!String.IsNullOrEmpty(tmp))
            {
                notes.Add(tmp);
            }

            return notes.ToArray();
        }


        protected string[] SplitNoteRow(string note)
        {
            string regPattern = @".{1," + limitNoteRow + "}([ ]|$)";
            String pattern = regPattern;
            MatchCollection matchList = Regex.Matches(note, pattern);
            return matchList.Cast<Match>().Select(match => match.Value).ToArray();
        }

        public bool IsGenerated()
        {
            return isGenerated;
        }

    }
}
