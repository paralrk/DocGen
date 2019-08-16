using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.EmptyDocuments
{
    class SpecEmptyDocument : EmptyDocument
    {
        public SpecEmptyDocument() : base("Спецификация")
        {

        }
        protected override void SetColumnsWidth()
        {
            base.SetColumnsWidth();
            Excel.Range range;
            Excel.Range column;
            range = sheet.Range["A1:B1"];
            column = range.EntireColumn;
            column.ColumnWidth = 3;
            range = sheet.Range["C1"];
            column = range.EntireColumn;
            column.ColumnWidth = 4;
            range = sheet.Range["D1"];
            column = range.EntireColumn;
            column.ColumnWidth = 30;
            range = sheet.Range["E1"];
            column = range.EntireColumn;
            column.ColumnWidth = 34;
            range = sheet.Range["F1"];
            column = range.EntireColumn;
            column.ColumnWidth = 5;
            range = sheet.Range["G1"];
            column = range.EntireColumn;
            column.ColumnWidth = 13.5;
        }

        protected override void FormatCells()
        {
            base.FormatCells();
            ((Excel.Range)sheet.Cells).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Columns[4]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)sheet.Columns[5]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 7]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Cells.Rows[1]).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]].Orientation = 90;
            ((Excel.Range)sheet.Cells[1, 6]).Orientation = 90;
        }

        protected override void FillTitle()
        {
            base.FillTitle();
            Excel.Range cells = (Excel.Range)sheet.Cells;
            ((Excel.Range)cells[1, 1]).Value2 = "Формат";
            ((Excel.Range)cells[1, 2]).Value2 = "Зона";
            ((Excel.Range)cells[1, 3]).Value2 = "Поз.";
            ((Excel.Range)cells[1, 4]).Value2 = "Обозначение";
            ((Excel.Range)cells[1, 5]).Value2 = "Наименование";
            ((Excel.Range)cells[1, 6]).Value2 = "Кол.";
            ((Excel.Range)cells[1, 7]).Value2 = "Приме-чание";
            // font sizes
            ((Excel.Range)cells[1, 1]).Font.Size = 11;            

            ((Excel.Range)sheet.Columns[3]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[4]).ShrinkToFit = false;
            ((Excel.Range)sheet.Columns[5]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[6]).ShrinkToFit = true;
            ((Excel.Range)sheet.Columns[7]).ShrinkToFit = true;
            ((Excel.Range)cells[1, 7]).WrapText = true;
        }
    }
}
