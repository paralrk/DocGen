using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.EmptyDocuments
{
    class PE3EmptyDocument : EmptyDocument
    {
        public PE3EmptyDocument() : base ("Перечень элементов")
        {

        }
        protected override void SetColumnsWidth()
        {
            base.SetColumnsWidth();
            Excel.Range range;
            Excel.Range column;
            range = sheet.Range["A1"];
            column = range.EntireColumn;
            column.ColumnWidth = 3;
            range = sheet.Range["B1"];
            column = range.EntireColumn;
            column.ColumnWidth = 11;
            range = sheet.Range["C1"];
            column = range.EntireColumn;
            column.ColumnWidth = 57;
            range = sheet.Range["D1"];
            column = range.EntireColumn;
            column.ColumnWidth = 5;
            range = sheet.Range["E1"];
            column = range.EntireColumn;
            column.ColumnWidth = 13.5;
        }

        public override void FormatCells()
        {
            base.FormatCells();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 5]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Cells.Rows[1]).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)sheet.Cells[1, 1]).Orientation = 90;

        }

        public override void InitFormatCells()
        {
            base.InitFormatCells();
            ((Excel.Range)sheet.Cells).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Columns[3]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)sheet.Cells[1, 3]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Cells.Rows[1]).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)sheet.Cells[1, 1]).Orientation = 90;

        }

        protected override void FillTitle()
        {
            base.FillTitle();
            Excel.Range cells = (Excel.Range)sheet.Cells;
            ((Excel.Range)cells[1, 1]).Value2 = "Зона";
            ((Excel.Range)cells[1, 2]).Value2 = "Поз. обозначение";
            ((Excel.Range)cells[1, 3]).Value2 = "Наименование";
            ((Excel.Range)cells[1, 4]).Value2 = "Кол.";
            ((Excel.Range)cells[1, 5]).Value2 = "Примечание";
            // font sizes
            ((Excel.Range)cells[1, 2]).Font.Size = 10;
            ((Excel.Range)cells[1, 2]).WrapText = true;
        }
    }
}
