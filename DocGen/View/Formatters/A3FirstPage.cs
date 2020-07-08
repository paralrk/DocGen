using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.View.Blank;
using System.Diagnostics;

namespace DocGen.View.Formatters
{
    abstract class A3FirstPage
    {
        protected Excel.Worksheet sheet;
        public int Height { get; } = 37;
        public int RowsCount { get; } = 23;

        protected int pageCount = 1;

        public A3FirstPage(Excel.Worksheet sheet, int pageCount)
        {
            this.sheet = sheet;
            this.pageCount = pageCount;
        }

        public void Format()
        {
            SetRowsHeight();
            SetColumnsWidth();
            MergeCells();
            DrawBorders();
            FillBlank();
            FillBlankText();
        }

        private void SetRowsHeight()
        {
            sheet.Range["A1:A2"].UnMerge();
            // insert additional rows

            sheet.Range["26:37"].Insert();

            Excel.Range range;

            range = sheet.Rows;
            range.RowHeight = 26; // 9mm

            range = sheet.Range["1:1"];
            range.RowHeight = 24;
            range = sheet.Range["2:2"];
            range.RowHeight = 50;

            range = sheet.Range["26:26"];
            range.RowHeight = 5;
            range = sheet.Range["27:27"];
            range.RowHeight = 37; // 6mm
            range = sheet.Range["28:28"];
            range.RowHeight = 23; // 8mm
            range = sheet.Range["29:37"];
            range.RowHeight = 15; // 5mm
        }
        virtual protected void SetColumnsWidth()
        {
            // set width for extra columns (added by subclasses)
            Excel.Range range;
            Excel.Range column;
            // separator can be ; or ,
            var sep = (string)Globals.ThisAddIn.Application.
                International[Excel.XlApplicationInternational.xlListSeparator];
            string str = "A1" + sep + "F1" + sep + "K1:L1" + sep + "O1" + sep + "S1:X1" + sep + "AM:AP";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 2;
            range = sheet.Range["B1:D1"];
            column = range.EntireColumn;
            column.ColumnWidth = 2.86;
            range = sheet.Range["E1"];
            column = range.EntireColumn;
            column.ColumnWidth = 1;
            str = "G1" + sep + "M1:N1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 3.57;
            str = "H1:I1" + sep + "Y1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 4;
            range = sheet.Range["J1"];
            column = range.EntireColumn;
            column.ColumnWidth = 2.57;
            str = "P1" + sep + "R1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 6;
            range = sheet.Range["Q1"];
            column = range.EntireColumn;
            column.ColumnWidth = 10;
            str = "Z1" + sep + "AI1" + sep + "AQ1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 4.57;
            range = sheet.Range["AA1"];
            column = range.EntireColumn;
            column.ColumnWidth = 12;
            range = sheet.Range["AB1"];
            column = range.EntireColumn;
            column.ColumnWidth = 3;
            range = sheet.Range["AC1"];
            column = range.EntireColumn;
            column.ColumnWidth = 4.43;
            range = sheet.Range["AD1"];
            column = range.EntireColumn;
            column.ColumnWidth = 5.29;
            str = "AE1" + sep + "AJ1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 5;
            str = "AF1" + sep + "AK1:AL1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 7;
            str = "AG1" + sep + "AS1:AT1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 4.43;
            range = sheet.Range["AH1"];
            column = range.EntireColumn;
            column.ColumnWidth = 6.57;
            range = sheet.Range["AR1"];
            column = range.EntireColumn;
            column.ColumnWidth = 1.29;
        }
        virtual protected void MergeCells()
        {
            // common for A4 and A3
            //left cells
            sheet.Range["A1:A6"].Merge();
            sheet.Range["B1:B6"].Merge();
            sheet.Range["A7:A13"].Merge();
            sheet.Range["B7:B13"].Merge();
            sheet.Range["A14:B15"].Merge();
            sheet.Range["A16:A19"].Merge();
            sheet.Range["B16:B19"].Merge();
            sheet.Range["A20:A22"].Merge();
            sheet.Range["B20:B22"].Merge();
            sheet.Range["A23:A26"].Merge();
            sheet.Range["B23:B26"].Merge();
            sheet.Range["A27:A31"].Merge();
            sheet.Range["B27:B31"].Merge();
            sheet.Range["A32:A36"].Merge();
            sheet.Range["B32:B36"].Merge();
            //bottom line
            sheet.Range["A37:AG37"].Merge();
            sheet.Range["AH37:AM37"].Merge();
            sheet.Range["AN37:AT37"].Merge();
            //blank
            sheet.Range["C26:AA36"].Merge();
            // доп. графы над основной надписью
            sheet.Range["AB26:AG28"].Merge();
            sheet.Range["AH26:AH27"].Merge();
            sheet.Range["AI26:AL27"].Merge();
            sheet.Range["AM26:AT27"].Merge();
            sheet.Range["AH28:AT28"].Merge();
            // № докум
            for (int i = 29; i <= 31; i++)
            {
                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
            }
            // обозначение
            sheet.Range["AH29:AT31"].Merge();
            // разработал, проверил, н. контр., утвердил
            for (int i = 32; i <= 36; i++)
            {
                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
            }
            // наименование
            sheet.Range["AH32:AM35"].Merge();
            // тип документа
            sheet.Range["AH36:AM36"].Merge();
            // литера
            sheet.Range["AN32:AP32"].Merge();
            // лист
            sheet.Range["AQ32:AR32"].Merge();
            sheet.Range["AQ33:AR33"].Merge();
            // листов
            sheet.Range["AS32:AT32"].Merge();
            sheet.Range["AS33:AT33"].Merge();
            // организация
            sheet.Range["AN34:AT36"].Merge();
        }
        virtual protected void DrawBorders()
        {
            // common for blank A3
            // clean
            sheet.Range["A1:AT37"].Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            // left cells
            sheet.Range["A1:B13"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["A16:B36"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // blank
            sheet.Range["AH26:AT31"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["AN32:AT36"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["AB29:AG36"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["AB29:AG30"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные
            sheet.Range["AB32:AG36"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные
            sheet.Range["AH36:AM36"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight =
                Excel.XlBorderWeight.xlMedium; // нижняя внешняя
            sheet.Range["C26:AA36"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight =
                Excel.XlBorderWeight.xlMedium; // нижняя внешняя
        }
        virtual protected void FillBlank()
        {
            // common for PE3 and Specification
            // vertical text orientation
            sheet.Range["A1:B36"].Orientation = 90;

            // left cells default text
            sheet.Range["A1:A6"].Value2 = "Перв. примен.";
            sheet.Range["A7:A13"].Value2 = "Справ. №";
            sheet.Range["A16:A19"].Value2 = "Подп. и дата";
            sheet.Range["A20:A22"].Value2 = "Инв. № дубл.";
            sheet.Range["A23:A26"].Value2 = "Взам. инв. №";
            sheet.Range["A27:A31"].Value2 = "Подп. и дата";
            sheet.Range["A32:A36"].Value2 = "Инв. № подл.";

            // blank default text
            sheet.Range["AB31"].Value2 = "Изм.";
            sheet.Range["AC31"].Value2 = "Лист";
            sheet.Range["AD31:AE31"].Value2 = "№ докум.";
            sheet.Range["AF31"].Value2 = "Подп.";
            sheet.Range["AG31"].Value2 = "Дата";
            sheet.Range["AB32:AC32"].Value2 = "Разраб.";
            sheet.Range["AB33:AC33"].Value2 = "Пров.";
            sheet.Range["AB35:AC35"].Value2 = "Н. контр.";
            sheet.Range["AB36:AC36"].Value2 = "Утв.";
            sheet.Range["AN32:AP32"].Value2 = "Лит.";
            sheet.Range["AQ32:AR32"].Value2 = "Лист";
            sheet.Range["AS32:AT32"].Value2 = "Листов";
            sheet.Range["AH37:AM37"].Value2 = "Копировал:";
            sheet.Range["AN37:AT37"].Value2 = "Формат А3";
            if (pageCount > 1)
            {
                sheet.Range["AQ33:AR33"].Value2 = 1;
            }
            sheet.Range["AS33:AT33"].Value2 = pageCount;

            // font sizes
            sheet.Range["A1:B36"].Font.Size = 11;
            sheet.Range["AB29:AF36"].Font.Size = 11;
            sheet.Range["AG29:AG36"].Font.Size = 8;
            sheet.Range["AG31:AG31"].Font.Size = 11;
            sheet.Range["A37:AT37"].Font.Size = 11;
            sheet.Range["AN32:AT33"].Font.Size = 11;
            sheet.Range["AN33:AP33"].Font.Size = 9;
            sheet.Range["AN34:AT36"].Font.Size = 12;
            sheet.Range["AH26:AT27"].Font.Size = 11;
            sheet.Range["AH28:AT31"].Font.Size = 20;
            sheet.Range["AH32:AM35"].Font.Size = 20;
            sheet.Range["AH36:AM36"].Font.Size = 12;

            // text align
            //HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["A1:B36"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["AB26:AT28"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["AH29:AT31"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["AH32:AT36"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["AB29:AG31"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["AB32:AE36"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["AF32:AG36"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["AH26:AM36"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["AH37:AM37"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["AN37:AT37"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            sheet.Range["A14:B15"].WrapText = true;
            sheet.Range["AH32:AM35"].WrapText = true;
        }

        virtual protected void FillBlankText()
        {
            A3FirstPageBlankTextFiller filler = new A3FirstPageBlankTextFiller();
            filler.Fill(sheet);
        }
    }
}
