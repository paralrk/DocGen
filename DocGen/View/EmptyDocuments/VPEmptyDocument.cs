using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.EmptyDocuments
{
    class VPEmptyDocument : EmptyDocument
    {
        public VPEmptyDocument() : base("Ведомость покупных изделий")
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
            column.ColumnWidth = 38;
            range = sheet.Range["C1"];
            column = range.EntireColumn;
            column.ColumnWidth = 20;
            range = sheet.Range["D1"];
            column = range.EntireColumn;
            column.ColumnWidth = 32;
            range = sheet.Range["E1"];
            column = range.EntireColumn;
            column.ColumnWidth = 27;
            range = sheet.Range["F1"];
            column = range.EntireColumn;
            column.ColumnWidth = 36;
            range = sheet.Range["G1:J1"];
            column = range.EntireColumn;
            column.ColumnWidth = 8;
            range = sheet.Range["K1"];
            column = range.EntireColumn;
            column.ColumnWidth = 12;

        }

        public override void FormatCells()
        {
            base.FormatCells();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 11]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Cells.Rows[1]).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)sheet.Cells[1, 1]).Orientation = 90;
        }

        public override void InitFormatCells()
        {
            base.InitFormatCells();
            ((Excel.Range)sheet.Cells).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)sheet.Columns[2]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //FormatCells();
        }

        protected override void FillTitle()
        {
            base.FillTitle();
            Excel.Range cells = (Excel.Range)sheet.Cells;
            ((Excel.Range)cells[1, 1]).Value2 = "№ строки";
            ((Excel.Range)cells[1, 2]).Value2 = "Наименование";
            ((Excel.Range)cells[1, 3]).Value2 = "Код продукции";
            ((Excel.Range)cells[1, 4]).Value2 = "Обозначение документа на поставку";
            ((Excel.Range)cells[1, 5]).Value2 = "Поставщик";
            ((Excel.Range)cells[1, 6]).Value2 = "Куда входит (обозначение)";
            ((Excel.Range)cells[1, 7]).Value2 = "на изделие";
            ((Excel.Range)cells[1, 8]).Value2 = "в комплекты";
            ((Excel.Range)cells[1, 9]).Value2 = "на регулир.";
            ((Excel.Range)cells[1, 10]).Value2 = "всего";
            ((Excel.Range)cells[1, 11]).Value2 = "Примечание";
            // font sizes
            ((Excel.Range)cells[1, 1]).Font.Size = 11;
            ((Excel.Range)cells[1, 7]).Font.Size = 11;
            ((Excel.Range)cells[1, 8]).Font.Size = 11;
            ((Excel.Range)cells[1, 9]).Font.Size = 11;
            ((Excel.Range)cells[1, 10]).Font.Size = 11;

            //((Excel.Range)sheet.Columns[3]).ShrinkToFit = true;
            ((Excel.Range)cells[1, 4]).WrapText = true;
            ((Excel.Range)cells[1, 6]).WrapText = true;
            ((Excel.Range)cells[1, 7]).WrapText = true;
            ((Excel.Range)cells[1, 8]).WrapText = true;
            ((Excel.Range)cells[1, 9]).WrapText = true;
            ((Excel.Range)cells[1, 10]).WrapText = true;
            ((Excel.Range)cells[1, 11]).WrapText = true;
        }
    }
}