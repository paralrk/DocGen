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
    class Specification : Document
    {
        //public List<RowSpec> Documents { get; private set; } = new List<RowSpec>();
        //private List<RowSpec> assemblies;
        //public List<RowSpec> Parts { get; private set; } = new List<RowSpec>();
        public List<RowSpec> Others { get; private set; } = new List<RowSpec>();
        //private List<RowSpec> materials;
        public List<RowSpec> NotesRows { get; private set; } = new List<RowSpec>();

        private int positionNumber = 5;
        private int positionInc = 2;
        private int maxNameLength = 28;


        public Specification() : base()
        {
            this.positionNumber = settings.StartPositionNumber;
            this.positionInc = settings.PositionInc;
            this.maxNameLength = settings.MaxNameLengthSpec;
            this.maxNoteLength = settings.MaxNoteLengthSpec;
            this.limitNoteRow = settings.LimitNoteRowSpec;
        }



        protected override void SortGroups()
        {
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
            Others = new List<RowSpec>();
            List<RowSpec> partRows;
            RowSpec row = new RowSpec();
            row.Name = "Прочие изделия";
            Others.Add(row);
            Others.Add(new RowSpec());
            string note = "";
            foreach (Model.Group group in componentList.Groups)
            {
                group.CountManufactures(settings.GroupLimitSpec);
                Components previous = new Components();
                foreach (Components c in group)
                {
                    partRows = new List<RowSpec>();
                    row = new RowSpec();
                    partRows.Add(row);
                    note = "";
                    row.Position = positionNumber.ToString();
                    positionNumber += positionInc;
                    string designators = DesignatorShortener.Short(c.GetDesignatorsList());

                    if (c.NoteNumber > 0 || c.NoteNumber1 > 0)
                    {
                        note = "Примеч. " + (c.NoteNumber > 0 ? c.NoteNumber.ToString() : "") +
                            (c.NoteNumber1 > 0 ? ", " + c.NoteNumber1.ToString() : "");
                        //row.Note = note;
                    }
                    if (!String.IsNullOrEmpty(note))
                    {
                        note += ", ";
                    }

                    // print header for subgroup and print Name and Note
                    if (group.IsInHeaders(c))
                    {
                        if (!c.IsEqualHeader(previous))
                        {
                            // empty row within groups
                            Others.Add(new RowSpec());
                            PrintHeader(c, group.GetHeader(c));
                        }
                        row = AddName(partRows, c.Part.ManufacturerPartNumber);
                        row = AddName(partRows, c.Part.Description);

                        row = AddNote(partRows, "", designators);
                    }
                    // print Name and Note without header  
                    else
                    {
                        if (group.IsInHeaders(previous))
                        {
                            // empty row within groups
                            Others.Add(new RowSpec());
                        }
                        row = AddName(partRows, c.Part.Type);
                        row = AddName(partRows, c.Part.ManufacturerPartNumber);
                        row = AddName(partRows, c.Part.Description);
                        row = AddName(partRows, c.Part.Manufacturer);

                        row = AddNote(partRows, note, designators);
                    }

                    row.Quantity = c.Quantity;
                    Others.AddRange(partRows);
                    previous = c;
                }
                // empty row within groups
                Others.Add(new RowSpec());
            }

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

        private void PrintHeader(Components c, GroupHeader header)
        {
            RowSpec row = new RowSpec();
            string note = "";
            row.Name = header.TypeDescription + " " + header.Manufacturer;
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
            DocumentsSection docs = new DocumentsSection();
            PartsSection parts = new PartsSection();
            List<IRow> rows = new List<IRow>();

            rows.AddRange(docs.GetDocuments());
            rows.Add(new RowSpec());
            rows.Add(new RowSpec());
            rows.AddRange(parts.GetDocuments());
            rows.Add(new RowSpec());
            rows.Add(new RowSpec());
            rows.AddRange(Others);
            rows.Add(new RowSpec());
            rows.Add(new RowSpec());
            rows.AddRange(NotesRows);
            return rows;
        }
    }
}
