using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocGen.Model;
using DocGen.Model.Documents;
using DocGen.Model.Documents.Comparers;
using DocGen.Model.Documents.Templates;
using DocGen.Utils;

namespace DocGen.Model.Documents
{
    class VP : AltiumDocument
    {

        public List<RowVP> Others { get; private set; } = new List<RowVP>();
        public List<RowVP> NotesRows { get; private set; } = new List<RowVP>();

        private int maxNameLength = 30;

        public VP() : base()
        {
            this.maxNameLength = settings.MaxNameLengthSpec;
            this.maxNoteLength = settings.MaxNoteLengthSpec;
            this.limitNoteRow = settings.LimitNoteRowSpec;
        }

        protected override void SortBOM()
        {
            Debug.WriteLine("Sorting BOM VP");
            base.SortBOM();
            Debug.WriteLine("Sorting BOM VP with VPComparer");
            bom.Sort(new VPComparer());
            Debug.WriteLine("BOM VP with VPComparer is sorted");
        }

        protected override void SortGroups()
        {
            Debug.WriteLine("Sorting Groups VP");
            foreach (Model.Group group in componentList.Groups)
            {
                group.CList = group.CList.OrderBy(c => c.GetDesignators().Length).
                    ThenBy(c => c.GetDesignators()).ToList();
                group.CList = group.CList.OrderBy(c => c.Part.Manufacturer).
                    ThenBy(c => c.Part.ManufacturerPartNumber).ToList();
            }
        }

        protected override void FillDocumentRows()
        {
            Others = new List<RowVP>();
            List<RowVP> partRows;
            RowVP row = new RowVP();
            Others.Add(row);
            string note = "";
            foreach (Model.Group group in componentList.Groups)
            {
                // to do - add settings to VP
                group.CountManufactures(settings.GroupLimitSpec);
                Components previous = new Components();
                foreach (Components c in group)
                {
                    partRows = new List<RowVP>();
                    row = new RowVP();
                    partRows.Add(row);
                    note = "";

                    if (c.NoteNumber > 0 || c.NoteNumber1 > 0)
                    {
                        note = "Примеч. " + (c.NoteNumber > 0 ? c.NoteNumber.ToString() : "") +
                            (c.NoteNumber1 > 0 ? ", " + c.NoteNumber1.ToString() : "");
                        //row.Note = note;
                    }

                    if (!String.IsNullOrEmpty(c.Part.Description))
                    {
                        c.Part.Description = "(" + c.Part.Description + ")";
                    }
                    // print header for subgroup and print Name
                    if (group.IsInHeaders(c))
                    {
                        if (!c.IsEqualHeader(previous))
                        {
                            // empty row within groups
                            Others.Add(new RowVP());
                            PrintHeader(c, group.GetHeader(c));
                        }
                        row = AddName(partRows, c.Part.ManufacturerPartNumber);
                        row = AddName(partRows, c.Part.Description);
                    }
                    // print Name without header  
                    else
                    {
                        if (group.IsInHeaders(previous))
                        {
                            // empty row within groups
                            Others.Add(new RowVP());
                        }
                        row.Designation = c.Part.Manufacturer;
                        row = AddName(partRows, c.Part.Type);
                        row = AddName(partRows, c.Part.ManufacturerPartNumber);
                        row = AddName(partRows, c.Part.Description);
                    }

                    row.Note = note;
                    row.QuantityProduct = c.Quantity;
                    row.QuantityTotal = c.Quantity;

                    Others.AddRange(partRows);
                    previous = c;
                    Others.Add(new RowVP());
                }
                // empty row within groups
                Others.Add(new RowVP());
            }

            // generating notes
            NotesRows.Add(new RowVP(0, "Примечания - ", "", "", "", "", 0, 0, 0, 0, ""));
            NotesRows.Add(new RowVP(0, "Допускается замена - ", "", "", "", "", 0, 0, 0, 0, ""));
            row = new RowVP();
            NotesRows.Add(row);

            for (int noteNumber = 0; noteNumber < notes.Count; noteNumber++)
            {
                string noteRow = (noteNumber + 1) + " " + notes[noteNumber];
                row = AddNoteRow(NotesRows, noteRow);
            }

        }

        private RowVP AddName(List<RowVP> partRows, string name)
        {
            RowVP row = partRows.Last();
            if (!String.IsNullOrEmpty(name))
            {
                if ((row.Name.Length + name.Length) < maxNameLength)
                {
                    if (!String.IsNullOrEmpty(row.Name))
                    {
                        row.Name += " ";
                    }
                    row.Name += name;
                }
                else
                {
                    row = new RowVP();
                    row.Name = name;
                    partRows.Add(row);
                }
            }
            return row;
        }

        private RowVP AddNoteRow(List<RowVP> notesRows, string noteRow)
        {
            RowVP row;
            // to do - load this number from settings
            limitNoteRow = 150;

            // split long string into several short
            string[] notesArray = SplitNoteRow(noteRow);

            for (int i = 0; i < notesArray.Length; i++)
            {
                row = new RowVP();
                row.Name = notesArray[i];
                notesRows.Add(row);
            }
            return notesRows.Last();
        }

        private void PrintHeader(Components c, GroupHeader header)
        {
            RowVP row = new RowVP();
            string note = "";
            row.Name = header.TypeDescription;
            row.Designation = header.Manufacturer;
            if (header.NoteNumber > 0 || header.NoteNumber1 > 0)
            {
                note = "Примеч. " + (header.NoteNumber > 0 ? header.NoteNumber.ToString() : "") +
                    (header.NoteNumber1 > 0 ? ", " + header.NoteNumber1.ToString() : "");
                row.Note = note;
            }
            Others.Add(row);
        }
        public override List<IRow> CombineRows()
        {
            List<IRow> rows = new List<IRow>();

            rows.AddRange(Others);
            rows.Add(new RowVP());
            rows.Add(new RowVP());
            rows.AddRange(NotesRows);
            return rows;
        }

    }
}
