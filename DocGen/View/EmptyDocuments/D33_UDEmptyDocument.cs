using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.EmptyDocuments
{
    class D33_UDEmptyDocument : EmptyDocument
    {
        public D33_UDEmptyDocument() : base("Д33-УД")
        {

        }
        protected override void SetColumnsWidth()
        {
            base.SetColumnsWidth();
            Excel.Range range;
            Excel.Range column;
            range = sheet.Range["A1"];
            column = range.EntireColumn;
            column.ColumnWidth = 22;
            range = sheet.Range["B1"];
            column = range.EntireColumn;
            column.ColumnWidth = 9;
            range = sheet.Range["C1"];
            column = range.EntireColumn;
            column.ColumnWidth = 9;
            range = sheet.Range["D1"];
            column = range.EntireColumn;
            column.ColumnWidth = 20;
            range = sheet.Range["E1"];
            column = range.EntireColumn;
            column.ColumnWidth = 16;
        }

        public override void FormatCells()
        {
            base.FormatCells();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 6]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Cells.Rows[1]).
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        public override void InitFormatCells()
        {
            base.InitFormatCells();
            ((Excel.Range)sheet.Cells).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Columns[1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            FormatCells();
        }

        protected override void FillTitle()
        {
            base.FillTitle();
            Excel.Range cells = (Excel.Range)sheet.Cells;

            ((Excel.Range)cells[1, 1]).Value2 = "Обозначение";
            ((Excel.Range)cells[1, 2]).Value2 = "Разработал";
            ((Excel.Range)cells[1, 3]).Value2 = "Изготовил";
            ((Excel.Range)cells[1, 4]).Value2 = "Согласовано";
            ((Excel.Range)cells[1, 5]).Value2 = "Утвердил";

            ((Excel.Range)sheet.Columns[1]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[2]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[3]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[4]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[5]).ShrinkToFit = true;

        }
    }
}
