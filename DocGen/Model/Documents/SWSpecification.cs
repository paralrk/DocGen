using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocGen.Model;
using DocGen.Model.Documents;
using DocGen.Model.Documents.Templates;
using DocGen.Utils;

namespace DocGen.Model.Documents
{
    class SWSpecification : SWDocument
    {
        public List<RowSpec> DocumentsRows { get; private set; } = new List<RowSpec>();
        public List<RowSpec> AssembliesRows { get; private set; } = new List<RowSpec>();
        public List<RowSpec> PartsRows { get; private set; } = new List<RowSpec>();
        public List<RowSpec> OthersRows { get; private set; } = new List<RowSpec>();
        public List<RowSpec> MaterialsRows { get; private set; } = new List<RowSpec>();
        public List<RowSpec> NotesRows { get; private set; } = new List<RowSpec>();

        private List<RowSpec> specRows = new List<RowSpec>();





        private int positionNumber = 1;
        private int positionInc = 2;
        private int maxNameLength = 28;


        public SWSpecification() : base()
        {
            //this.positionNumber = settings.StartPositionNumber;
            this.positionInc = settings.PositionInc;
            this.maxNameLength = settings.MaxNameLengthSpec;
            this.maxNoteLength = settings.MaxNoteLengthSpec;
            this.limitNoteRow = settings.LimitNoteRowSpec;
        }



        protected override void FillDocumentRows(List<SWComponents> group, String section)
        {
            List<RowSpec> partRows;
            RowSpec row = new RowSpec();
            row.Name = section;
            specRows.Add(row);
            specRows.Add(new RowSpec());
            string note = "";
            foreach (SWComponents c in group)
            {
                partRows = new List<RowSpec>();
                row = new RowSpec();
                partRows.Add(row);
                note = "";
                row.Position = positionNumber.ToString();
                positionNumber += positionInc;

                if (c.ReplacementNumber > 0)
                {
                    note = "Примеч. " + c.ReplacementNumber.ToString();
                }
                if (!String.IsNullOrEmpty(note) && !String.IsNullOrEmpty(c.Note))
                {
                    note += ", ";
                }

                row.Designation = c.Designation;
                row.Name = c.Name;
                //row = AddName(partRows, c.Name);
                row = AddNote(partRows, note, c.Note);

                row.Quantity = c.Quantity;

                specRows.AddRange(partRows);
                specRows.Add(new RowSpec());
            }

            // empty row within groups
            specRows.Add(new RowSpec());

        }

        protected override void GenerateNotes()
        {
            RowSpec row = new RowSpec();
            // generating notes
            NotesRows.Add(new RowSpec("", "", "", "Примечания - ", "", 0, ""));
            NotesRows.Add(new RowSpec("", "", "", "Допускается замена - ", "", 0, ""));
            row = new RowSpec();
            NotesRows.Add(row);

            for (int noteNumber = 0; noteNumber < notes.Count; noteNumber++)
            {
                string noteRow = (noteNumber + 1) + " " + notes[noteNumber];
                row = AddNoteRow(NotesRows, noteRow);
            }
        }

        private RowSpec AddName(List<RowSpec> partRows, string name)
        {
            RowSpec row = partRows.Last();
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
                    row = new RowSpec();
                    row.Name = name;
                    partRows.Add(row);
                }
            }
            return row;
        }

        private RowSpec AddNote(List<RowSpec> partRows, string note, string designators)
        {
            RowSpec row = partRows.Last();

            // split long string into several short
            // example - "DA1, DA5, DA6, DA8" -> "DA1, DA5", "DA6, DA8"
            string[] desArray = SplitNote(designators);

            // if necessary, add new rows
            int noteRowsCount = 0;
            if (!String.IsNullOrEmpty(note))
            {
                noteRowsCount = 1;
            }

            while (partRows.Count < desArray.Length + noteRowsCount)
            {
                partRows.Add(new RowSpec());
            }

            int j = partRows.Count - desArray.Length;
            for (int i = 0; i < desArray.Length; i++)
            {
                partRows[j + i].Note = desArray[i];
            }

            if (!String.IsNullOrEmpty(note))
            {
                partRows[j - 1].Note = note;
            }
            return partRows.Last();
        }

        private RowSpec AddNoteRow(List<RowSpec> notesRows, string noteRow)
        {
            RowSpec row;

            // split long string into several short
            string[] notesArray = SplitNoteRow(noteRow);

            for (int i = 0; i < notesArray.Length; i++)
            {
                row = new RowSpec();
                row.Designation = notesArray[i];
                notesRows.Add(row);
            }
            return notesRows.Last();
        }


        public override List<IRow> CombineRows()
        {
            DocumentsSection docs = new DocumentsSection();
            List<IRow> rows = new List<IRow>();
            rows.AddRange(docs.GetDocuments());
            rows.Add(new RowSpec());
            rows.Add(new RowSpec());
            rows.AddRange(specRows);
            rows.Add(new RowSpec());
            rows.Add(new RowSpec());
            rows.AddRange(NotesRows);
            return rows;
        }
    }
}
