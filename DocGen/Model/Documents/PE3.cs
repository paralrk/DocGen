using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DocGen.Model;
using DocGen.Utils;
using DocGen.View.EmptyDocuments;
using DocGen.Model.Documents;

namespace DocGen.Model.Documents
{
    class PE3 : AltiumDocument
    {
        private List<RowPE3> rows = new List<RowPE3>();

        private int maxNameLength = 50;

        public PE3() : base()
        {
            this.maxNameLength = settings.MaxNameLengthPE3;
            this.limitNoteRow = settings.LimitNoteRowPE3;
        }

        protected override void SortGroups()
        {
            foreach (Group group in componentList.Groups)
            {
                group.CList = group.CList.OrderBy(
                    c => c.GetDesignators().Length).
                    ThenBy(c => c.GetDesignators()).ToList();
            }
        }

        protected override void FillDocumentRows()
        {
            rows = new List<RowPE3>();
            RowPE3 row;
            string note = "";
            foreach (Group group in componentList.Groups)
            {
                group.CountManufactures(settings.GroupLimitPE3);
                foreach (GroupHeader header in group.Headers)
                {
                    row = new RowPE3();
                    row.Name = header.TypeDescription + " " + header.Manufacturer;
                    if (header.NoteNumber > 0 || header.NoteNumber1 > 0)
                    {
                        note = "Примеч. " + (header.NoteNumber > 0 ? header.NoteNumber.ToString() : "") +
                            (header.NoteNumber1 > 0 ? ", " + header.NoteNumber1.ToString() : "");
                        row.Note = note;
                    }
                    rows.Add(row);
                }
                foreach (Components c in group)
                {
                    row = new RowPE3();
                    rows.Add(row);
                    row.Designator = DesignatorShortener.Short(c.GetDesignatorsList());
                    string description = null;
                    if (!String.IsNullOrEmpty(c.Part.Description))
                    {
                        description = "(" + c.Part.Description + ")";
                    }
                    if (group.IsInHeaders(c))
                    {
                        row = AddName(rows, c.Part.ManufacturerPartNumber);
                        row = AddName(rows, description);
                        //row.Name = c.Part.ManufacturerPartNumber + " " + c.Part.Description;
                    } else
                    {
                        row = AddName(rows, c.Part.Type);
                        row = AddName(rows, c.Part.ManufacturerPartNumber);
                        row = AddName(rows, description);
                        row = AddName(rows, c.Part.Manufacturer);
                        //row.Name = c.Part.Type + " " + c.Part.ManufacturerPartNumber +
                        //    " " + c.Part.Description + " " + c.Part.Manufacturer;
                    }
                    row.Quantity = c.Quantity;
                    if (c.NoteNumber > 0 || c.NoteNumber1 > 0)
                    {
                        note = "Примеч. " + (c.NoteNumber > 0 ? c.NoteNumber.ToString() : "") +
                            (c.NoteNumber1 > 0 ? ", " + c.NoteNumber1.ToString() : "");
                        row.Note = note;
                    }
                    //rows.Add(row);
                }
                // empty row within groups
                rows.Add(new RowPE3());
            }

            string noteRow = "";
            for (int noteNumber = 0; noteNumber < notes.Count; noteNumber++)
            {
                row = new RowPE3();
                noteRow = (noteNumber + 1) + " " + notes[noteNumber];
                row = AddNoteRow(rows, noteRow);
            }

        }

        private RowPE3 AddName(List<RowPE3> rows, string name)
        {
            RowPE3 row = rows.Last();
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
                    row = new RowPE3();
                    row.Name = name;
                    rows.Add(row);
                }
            }
            return row;
        }

        private RowPE3 AddNoteRow(List<RowPE3> notesRows, string noteRow)
        {
            RowPE3 row;

            // split long string into several short
            string[] notesArray = SplitNoteRow(noteRow);

            for (int i = 0; i < notesArray.Length; i++)
            {
                row = new RowPE3();
                row.Name = notesArray[i];
                notesRows.Add(row);
            }
            return notesRows.Last();
        }

        public override List<IRow> CombineRows()
        {
            List<IRow> exportingRows = new List<IRow>();
            exportingRows.AddRange(rows);
            return exportingRows;
        }
    }
}
