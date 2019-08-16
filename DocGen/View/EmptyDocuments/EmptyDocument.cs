using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.Utils;
using DocGen.View.Blank;

namespace DocGen.View.EmptyDocuments
{
    abstract class EmptyDocument
    {

        protected Excel.Worksheet sheet;
        string documentType;

        public EmptyDocument(string documentType) : base()
        {
            this.sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            this.documentType = documentType;

            Excel.Application xlApp = (Excel.Application)Globals.ThisAddIn.Application;
            xlApp.ActiveWindow.DisplayZeros = false;
        }

        public void NewDocument()
        {
            sheet = SheetHelper.getSheet();
            ListPage.AddDocument(sheet.Name, documentType);
            sheet.Activate();
        }

        public void Format()
        {
            ExcelHelper.DisableUpdating();
            SetRowsHeight();
            SetColumnsWidth();
            FormatCells();
            FillTitle();
            ExcelHelper.EnableUpdating();
        }

        protected virtual void SetRowsHeight()
        {
            Excel.Range range;

            sheet.Rows.RowHeight = 24.9; // 9mm

            // insert first row
            sheet.Range["1:1"].Insert();

            range = sheet.Range["1:1"];
            range.RowHeight = 41.75; // 15mm
            ((Excel.Range)sheet.Rows[1]).RowHeight = 43;
        }

        protected virtual void SetColumnsWidth()
        {

        }

        protected virtual void FormatCells()
        {
            Excel.Style style = StyleHelper.getDocGenMainStyle();
            sheet.Cells.Style = style;
        }

        protected virtual void FillTitle()
        {

        }

        public Excel.Worksheet getSheet()
        {
            return sheet;
        }
    }
}
