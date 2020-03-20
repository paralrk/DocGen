using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

using DocGen.View.EmptyDocuments;
using DocGen.Utils;

namespace DocGen.View.Unformatters
{
    abstract class Unformatter
    {
        protected Excel.Worksheet sheet;

        public Unformatter()
        {
            this.sheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
        }

        public void Unformat()
        {
            ExcelHelper.DisableUpdating();

            Stopwatch sw = new Stopwatch();
            Debug.WriteLine("Unformatting document");

            sw.Start();
            DeleteRows();
            sw.Stop();
            Debug.WriteLine("DeleteRows() Elapsed={0}", sw.Elapsed);

            sw.Start();
            ClearFormat();
            sw.Stop();
            Debug.WriteLine("ClearFormat() Elapsed={0}", sw.Elapsed);

            sw.Start();
            UnmergeCells();
            sw.Stop();
            Debug.WriteLine("UnmergeCells() Elapsed={0}", sw.Elapsed);

            sw.Start();
            DeleteColumns();
            sw.Stop();
            Debug.WriteLine("DeleteColumns() Elapsed={0}", sw.Elapsed);

            sw.Start();
            FormatEmptyDocument();
            sw.Stop();
            Debug.WriteLine("FormatEmptyDocument() Elapsed={0}", sw.Elapsed);

            ExcelHelper.EnableUpdating();
        }

        virtual protected void DeleteRows()
        {

        }

        virtual protected void ClearFormat()
        {
            sheet.Cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            // sheet.Cells.Style = "Normal";
        }

        virtual protected void UnmergeCells()
        {
            sheet.Cells.UnMerge();
        }

        virtual protected void DeleteColumns()
        {

        }

        virtual protected void FormatEmptyDocument()
        {
            EmptyDocumentsFactory factory = new EmptyDocumentsFactory();
            EmptyDocument emptyDocument = factory.GetEmptyDocument(sheet);
            if (emptyDocument != null)
            {
                emptyDocument.Format();
            }
        }

    }
}
