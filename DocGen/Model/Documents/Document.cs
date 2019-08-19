using System;
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
    abstract class Document
    {
        private List<Components> bom;
        protected ComponentList componentList;
        protected List<string> notes;
        protected Settings settings;
        public bool IsGenerated { get; private set; } = false;

        protected int maxNoteLength = 14;
        protected int limitNoteRow = 60;

        BOMReader bomReader;
        GroupFiller filler;

        public Document ()
        {
            SettingsFactory factory = new SettingsFactory();
            this.settings = factory.GetSettings();
        }

        public void Generate()
        {
            ReadBOM();
            if (bom != null)
            {
                SortBOM();
                FillGroups();
                SortGroups();
                CombineComponents();
                CreateNotes();
                FillDocumentRows();
                CombineRows();
                IsGenerated = true;
            }
        }

        private void ReadBOM()
        {
            bomReader = new BOMReader();
            bom = bomReader.ReadBOM();
        }
        private void SortBOM()
        {
            bom.Sort(new DesignatorComparer());
        }

        private void FillGroups()
        {
            filler = new GroupFiller();
            componentList = new ComponentList();
            componentList.Groups = filler.FillGroups(bom);
        }

        protected virtual void SortGroups()
        {
        }

        protected virtual void CombineComponents()
        {
            foreach (Model.Group group in componentList.Groups)
            {
                List<Components> combined = new List<Components>();
                Components last = null;
                foreach (Components c in group.CList)
                {
                    if (last == null)
                    {
                        last = c;
                    }
                    else if (last.Part.ManufacturerPartNumber.
                      Equals(c.Part.ManufacturerPartNumber))
                    {
                        last.AddComponent(c);
                    }
                    else
                    {
                        combined.Add(last);
                        last = c;
                    }
                }
                combined.Add(last);
                group.CList = combined;
            }
        }

        private void CreateNotes()
        {
            notes = new List<string>();
            foreach (Model.Group group in componentList.Groups)
            {
                foreach (Components c in group.CList)
                {
                    if (!String.IsNullOrEmpty(c.Part.Note))
                    {
                        if (!notes.Contains(c.Part.Note))
                        {
                            notes.Add(c.Part.Note);
                        }
                        c.NoteNumber = notes.IndexOf(c.Part.Note) + 1;
                    }

                    if (!String.IsNullOrEmpty(c.Part.Note1))
                    {
                        if (!notes.Contains(c.Part.Note1))
                        {
                            notes.Add(c.Part.Note1);
                        }
                        c.NoteNumber1 = notes.IndexOf(c.Part.Note1) + 1;
                    }
                }
            }
        }

        protected virtual void FillDocumentRows()
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

        internal ComponentList ComponentList
        {
            get => default;
            set
            {
            }
        }
    }
}
