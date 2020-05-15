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
            sheet.Range["1:1"].Insert();
            sheet.Range["A1:A2"].Merge();
            sheet.Range["B1:B2"].Merge();
            sheet.Range["C1:C2"].Merge();
            sheet.Range["D1:D2"].Merge();
            sheet.Range["E1:E2"].Merge();
            sheet.Range["F1:F2"].Merge();
            sheet.Range["G1:J1"].Merge();
            sheet.Range["K1:K2"].Merge();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 11]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 11]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
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
            sheet.Range["A1:A2"].Value2 = "№ строки";
            sheet.Range["B1:B2"].Value2 = "Наименование";
            sheet.Range["C1:C2"].Value2 = "Код продукции";
            sheet.Range["D1:D2"].Value2 = "Обозначение документа на поставку";
            sheet.Range["E1:E2"].Value2 = "Поставщик";
            sheet.Range["F1:F2"].Value2 = "Куда входит (обозначение)";
            sheet.Range["G1:J1"].Value2 = "Количество";
            ((Excel.Range)cells[2, 7]).Value2 = "на из- делие";
            ((Excel.Range)cells[2, 8]).Value2 = "в ком- плекты";
            ((Excel.Range)cells[2, 9]).Value2 = "на ре- гулир.";
            ((Excel.Range)cells[2, 10]).Value2 = "всего";
            sheet.Range["K1:K2"].Value2 = "Приме- чание";
            // font sizes
            sheet.Range["A1:A2"].Font.Size = 11;
            ((Excel.Range)cells[2, 7]).Font.Size = 11;
            ((Excel.Range)cells[2, 8]).Font.Size = 11;
            ((Excel.Range)cells[2, 9]).Font.Size = 11;
            ((Excel.Range)cells[2, 10]).Font.Size = 11;

            //((Excel.Range)sheet.Columns[3]).ShrinkToFit = true;
            sheet.Range["D1:D2"].WrapText = true;
            sheet.Range["F1:F2"].WrapText = true;
            ((Excel.Range)cells[2, 7]).WrapText = true;
            ((Excel.Range)cells[2, 8]).WrapText = true;
            ((Excel.Range)cells[2, 9]).WrapText = true;
            ((Excel.Range)cells[2, 10]).WrapText = true;
            sheet.Range["K1:K2"].WrapText = true;
        }
    }
}