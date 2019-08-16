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
    abstract class A4FirstPage
    {
        protected Excel.Worksheet sheet;
        public int Height { get; } = 37;
        public int RowsCount { get; } = 23;

        protected int pageCount = 1;

        public A4FirstPage(Excel.Worksheet sheet, int pageCount)
        {
            this.sheet = sheet;
            this.pageCount = pageCount;
        }

        public void format()
        {
            Stopwatch sw = new Stopwatch();
            Debug.WriteLine("Formatting Spec First Page document");

            sw.Start();
            SetRowsHeight();
            sw.Stop();
            Debug.WriteLine("SetRowsHeight() Elapsed={0}", sw.Elapsed);

            sw.Start();
            SetColumnsWidth();
            sw.Stop();
            Debug.WriteLine("SetColumnsWidth() Elapsed={0}", sw.Elapsed);

            sw.Start();
            MergeCells();
            sw.Stop();
            Debug.WriteLine("MergeCells() Elapsed={0}", sw.Elapsed);

            sw.Start();
            DrawBorders();
            sw.Stop();
            Debug.WriteLine("DrawBorders() Elapsed={0}", sw.Elapsed);

            sw.Start();
            FillBlank();
            sw.Stop();
            Debug.WriteLine("FillBlank() Elapsed={0}", sw.Elapsed);

            sw.Start();
            FillBlankText();
            sw.Stop();
            Debug.WriteLine("FillBlankText() Elapsed={0}", sw.Elapsed);

        }

        private void SetRowsHeight()
        {
            
            
            // insert additional rows

            Debug.WriteLine("Adding rows A4FirstPage");
            Stopwatch sw = new Stopwatch();
            sw.Start();
            sheet.Range["25:37"].Insert();
            sw.Stop();
            Debug.WriteLine("Rows added. Elapsed={0}", sw.Elapsed);

            Debug.WriteLine("Set rows height A4FirstPage");
            sw.Reset();
            sw.Start();

            Excel.Range range;

            range = sheet.Rows;
            range.RowHeight = 25; // 9mm

            range = sheet.Range["1:1"];
            range.RowHeight = 41.75; // 15mm
            //range = sheet.Range["2:24"];
            //range.RowHeight = 24.9; // 9mm
            range = sheet.Range["25:25"];
            range.RowHeight = 8; // 3mm
            range = sheet.Range["26:26"];
            range.RowHeight = 17; // 6mm
            range = sheet.Range["27:28"];
            range.RowHeight = 23; // 8mm
            range = sheet.Range["29:37"];
            range.RowHeight = 14.5; // 5mm

            sw.Stop();
            Debug.WriteLine("Rows height set. Elapsed={0}", sw.Elapsed);
        }
        virtual protected void SetColumnsWidth()
        {
            // set width for extra columns (added by subclasses)
            Excel.Range range;
            Excel.Range column;
            // separator can be ; or ,
            var sep = (string)Globals.ThisAddIn.Application.
                International[Excel.XlApplicationInternational.xlListSeparator];
            string str = "A1" + sep + "F1" + sep + "K1:L1" + sep + "O1" + sep + "S1:X1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 1.7; // 5mm
            str = "B1:D1" + sep + "M1:N1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 2.5; // 7mm
            range = sheet.Range["E1"];
            column = range.EntireColumn;
            column.ColumnWidth = 0.92; // 3mm
            range = sheet.Range["G1:I1"];
            column = range.EntireColumn;
            column.ColumnWidth = 3.5; // 9mm
            range = sheet.Range["J1"];
            column = range.EntireColumn;
            column.ColumnWidth = 2; // 6mm
            str = "P1" + sep + "R1";
            range = sheet.Range[str];
            column = range.EntireColumn;
            column.ColumnWidth = 5.9; // 14mm
            range = sheet.Range["Q1"];
            column = range.EntireColumn;
            column.ColumnWidth = 10; // 23mm
            range = sheet.Range["Y1:Z1"];
            column = range.EntireColumn;
            column.ColumnWidth = 3.9; //10 mm
        }
        virtual protected void MergeCells()
        {
            // common for PE3 and Specification
            //left cells
            sheet.Range["A1:A6"].Merge();
            sheet.Range["B1:B6"].Merge();
            sheet.Range["A7:A13"].Merge();
            sheet.Range["B7:B13"].Merge();
            sheet.Range["A14:B15"].Merge();
            sheet.Range["B7:B13"].Merge();
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
            sheet.Range["A37:L37"].Merge();
            sheet.Range["M37:R37"].Merge();
            sheet.Range["S37:Z37"].Merge();
            //blank
            sheet.Range["C25:Z25"].Merge();
            // доп. графы над основной надписью
            sheet.Range["C26:L28"].Merge();
            sheet.Range["M26:N27"].Merge();
            sheet.Range["O26:R27"].Merge();
            sheet.Range["S26:Z27"].Merge();
            sheet.Range["M28:Z28"].Merge();
            // лист, № докум, подп, дата
            for (int i = 29; i <= 31; i++)
            {
                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 8]].Merge();
                sheet.Range[sheet.Cells[i, 9], sheet.Cells[i, 10]].Merge();
                sheet.Range[sheet.Cells[i, 11], sheet.Cells[i, 12]].Merge();
            }
            // обозначение
            sheet.Range["M29:Z31"].Merge();
            // разработал, проверил, н. контр., утвердил
            for (int i = 32; i <= 36; i++)
            {
                sheet.Range[sheet.Cells[i, 3], sheet.Cells[i, 5]].Merge();
                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 8]].Merge();
                sheet.Range[sheet.Cells[i, 9], sheet.Cells[i, 10]].Merge();
                sheet.Range[sheet.Cells[i, 11], sheet.Cells[i, 12]].Merge();
            }
            // наименование
            sheet.Range["M32:R35"].Merge();
            // тип документа
            sheet.Range["M36:R36"].Merge();
            // литера
            sheet.Range["S32:U32"].Merge();
            // лист
            sheet.Range["V32:X32"].Merge();
            sheet.Range["V33:X33"].Merge();
            // листов
            sheet.Range["Y32:Z32"].Merge();
            sheet.Range["Y33:Z33"].Merge();
            // организация
            sheet.Range["S34:Z36"].Merge();
        }
        virtual protected void DrawBorders()
        {
            // common for PE3 and Specification
            sheet.Range["A1:Z37"].Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            // left cells
            sheet.Range["A1:B13"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["A16:B36"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // blank
            sheet.Range["M26:Z31"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["S32:Z36"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["C29:L36"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["C29:L30"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные
            sheet.Range["C32:L36"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные
            sheet.Range["C25:Z25"].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight =
                Excel.XlBorderWeight.xlMedium; // правая внешняя
            sheet.Range["M36:R36"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight =
                Excel.XlBorderWeight.xlMedium; // правая внешняя
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
            sheet.Range["C31"].Value2 = "Изм.";
            sheet.Range["D31:E31"].Value2 = "Лист";
            sheet.Range["F31:H31"].Value2 = "№ докум.";
            sheet.Range["I31:J31"].Value2 = "Подп.";
            sheet.Range["K31:L31"].Value2 = "Дата";
            sheet.Range["C32:E32"].Value2 = "Разраб.";
            sheet.Range["C33:E33"].Value2 = "Пров.";
            sheet.Range["C35:E35"].Value2 = "Н. контр.";
            sheet.Range["C36:E36"].Value2 = "Утв.";
            sheet.Range["S32:U32"].Value2 = "Лит.";
            sheet.Range["V32:X32"].Value2 = "Лист";
            sheet.Range["Y32:Z32"].Value2 = "Листов";
            sheet.Range["M37:R37"].Value2 = "Копировал:";
            sheet.Range["S37:Z37"].Value2 = "Формат А4";
            if (pageCount > 1)
            {
                sheet.Range["V33:X33"].Value2 = 1;
            }
            sheet.Range["Y33:Z33"].Value2 = pageCount;

            // font sizes
            sheet.Range["A1:B36"].Font.Size = 11;
            sheet.Range["C29:J36"].Font.Size = 11;
            sheet.Range["K29:L36"].Font.Size = 8;
            sheet.Range["K31:L31"].Font.Size = 11;
            sheet.Range["A37:Z37"].Font.Size = 11;
            sheet.Range["S32:Z33"].Font.Size = 11;
            sheet.Range["S33:U33"].Font.Size = 9;
            sheet.Range["S34:Z36"].Font.Size = 12;
            sheet.Range["M26:Z27"].Font.Size = 11;
            sheet.Range["M28:Z31"].Font.Size = 20;
            sheet.Range["M32:R35"].Font.Size = 14;
            sheet.Range["M32:R35"].Font.Size = 20;
            sheet.Range["M36:R36"].Font.Size = 12;

            // text align
            //HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["A1:B36"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["C26:Z28"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["M29:Z31"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["M32:Z36"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["C29:L31"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["C32:H36"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["I32:L36"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["M26:Z36"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["M37:R37"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["S37:Z37"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
        }

        virtual protected void FillBlankText()
        {
            A4FirstPageBlankTextFiller filler = new A4FirstPageBlankTextFiller();
            filler.Fill(sheet);
        }
    }
}
