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
    abstract class A4SecondPage
    {
        protected Excel.Worksheet sheet;
        public int Height { get; } = 35;
        public int RowsCount { get; } = 29;
        protected int pageNumber = 2;
        protected int firstRow;

        public A4SecondPage(Excel.Worksheet sheet, int firstRow, int pageNumber)
        {
            this.sheet = sheet;
            this.firstRow = firstRow;
            this.pageNumber = pageNumber;
        }

        public void format()
        {
            Stopwatch sw = new Stopwatch();
            Debug.WriteLine("Formatting Spec Second Page document");

            sw.Start();
            SetRowsHeight();
            sw.Stop();
            Debug.WriteLine("SetRowsHeight() Elapsed={0}", sw.Elapsed);

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

        virtual protected void SetRowsHeight()
        {
            // separator can be ; or ,
            var sep = (string)Globals.ThisAddIn.Application.
                International[Excel.XlApplicationInternational.xlListSeparator];
            // insert additional rows
            string rowsRange = firstRow + ":" + firstRow + sep +
                    (firstRow + 29) + ":" + (firstRow + 33);
            sheet.Range[rowsRange].Insert();

            string str;
            Excel.Range range;
            str = firstRow + ":" + firstRow;
            range = sheet.Range[str];
            range.RowHeight = 37; // 15mm
            //str = (firstRow + 1) + ":" + (firstRow + 28);
            //range = sheet.Range[str];
            //range.RowHeight = 24.9; // 9mm
            str = (firstRow + 30) + ":" + (firstRow + 30) + sep +
                    (firstRow + 33) + ":" + (firstRow + 34);
            range = sheet.Range[str];
            range.RowHeight = 14.5; // 5mm
            str = (firstRow + 31) + ":" + (firstRow + 31);
            range = sheet.Range[str];
            range.RowHeight = 5.5;
            str = (firstRow + 32) + ":" + (firstRow + 32);
            range = sheet.Range[str];
            range.RowHeight = 9;

        }

        virtual protected void MergeCells()
        {
            // common for PE3 and Specification
            //left cells
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 14, 2]].Merge();
            // подпись и дата
            sheet.Range[sheet.Cells[firstRow + 15, 1], sheet.Cells[firstRow + 18, 1]].Merge();
            sheet.Range[sheet.Cells[firstRow + 15, 2], sheet.Cells[firstRow + 18, 2]].Merge();
            // Инв. № дубл.
            sheet.Range[sheet.Cells[firstRow + 19, 1], sheet.Cells[firstRow + 21, 1]].Merge();
            sheet.Range[sheet.Cells[firstRow + 19, 2], sheet.Cells[firstRow + 21, 2]].Merge();
            // Взам. инв. №
            sheet.Range[sheet.Cells[firstRow + 22, 1], sheet.Cells[firstRow + 24, 1]].Merge();
            sheet.Range[sheet.Cells[firstRow + 22, 2], sheet.Cells[firstRow + 24, 2]].Merge();
            // Подп. и дата
            sheet.Range[sheet.Cells[firstRow + 25, 1], sheet.Cells[firstRow + 28, 1]].Merge();
            sheet.Range[sheet.Cells[firstRow + 25, 2], sheet.Cells[firstRow + 28, 2]].Merge();
            // Инв. № подл.
            sheet.Range[sheet.Cells[firstRow + 29, 1], sheet.Cells[firstRow + 33, 1]].Merge();
            sheet.Range[sheet.Cells[firstRow + 29, 2], sheet.Cells[firstRow + 33, 2]].Merge();

            //bottom line
            sheet.Range[sheet.Cells[firstRow + 34, 1], sheet.Cells[firstRow + 34, 12]].Merge();
            sheet.Range[sheet.Cells[firstRow + 34, 13], sheet.Cells[firstRow + 34, 18]].Merge();
            sheet.Range[sheet.Cells[firstRow + 34, 19], sheet.Cells[firstRow + 34, 26]].Merge();
            //blank

            // лист, № докум, подп, дата
            sheet.Range[sheet.Cells[firstRow + 31, 3], sheet.Cells[firstRow + 32, 3]].Merge();

            sheet.Range[sheet.Cells[firstRow + 30, 4], sheet.Cells[firstRow + 30, 5]].Merge();
            sheet.Range[sheet.Cells[firstRow + 30, 6], sheet.Cells[firstRow + 30, 8]].Merge();
            sheet.Range[sheet.Cells[firstRow + 30, 9], sheet.Cells[firstRow + 30, 10]].Merge();
            sheet.Range[sheet.Cells[firstRow + 30, 11], sheet.Cells[firstRow + 30, 12]].Merge();

            sheet.Range[sheet.Cells[firstRow + 31, 4], sheet.Cells[firstRow + 32, 5]].Merge();
            sheet.Range[sheet.Cells[firstRow + 31, 6], sheet.Cells[firstRow + 32, 8]].Merge();
            sheet.Range[sheet.Cells[firstRow + 31, 9], sheet.Cells[firstRow + 32, 10]].Merge();
            sheet.Range[sheet.Cells[firstRow + 31, 11], sheet.Cells[firstRow + 32, 12]].Merge();

            sheet.Range[sheet.Cells[firstRow + 33, 4], sheet.Cells[firstRow + 33, 5]].Merge();
            sheet.Range[sheet.Cells[firstRow + 33, 6], sheet.Cells[firstRow + 33, 8]].Merge();
            sheet.Range[sheet.Cells[firstRow + 33, 9], sheet.Cells[firstRow + 33, 10]].Merge();
            sheet.Range[sheet.Cells[firstRow + 33, 11], sheet.Cells[firstRow + 33, 12]].Merge();

            // обозначение
            sheet.Range[sheet.Cells[firstRow + 30, 13], sheet.Cells[firstRow + 33, 25]].Merge();
            // лист
            sheet.Range[sheet.Cells[firstRow + 30, 26], sheet.Cells[firstRow + 31, 26]].Merge();
            // номер листа
            sheet.Range[sheet.Cells[firstRow + 32, 26], sheet.Cells[firstRow + 33, 26]].Merge();
        }

        virtual protected void DrawBorders()
        {
            // common for PE3 and Specification
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 34, 26]].
                Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //left cells
            sheet.Range[sheet.Cells[firstRow + 15, 1], sheet.Cells[firstRow + 33, 2]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;

            // blank
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 26]].
                            Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 12]].
                            Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;  // внутренние горизонтальные

        }

        virtual protected void FillBlank()
        {
            // common for PE3 and Specification
            // vertical text orientation
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 33, 2]].
                Orientation = 90;

            // left cells default text
            sheet.Range[sheet.Cells[firstRow + 15, 1], sheet.Cells[firstRow + 18, 1]].
                Value2 = "Подп. и дата";
            sheet.Range[sheet.Cells[firstRow + 19, 1], sheet.Cells[firstRow + 21, 1]].
                Value2 = "Инв. № дубл.";
            sheet.Range[sheet.Cells[firstRow + 22, 1], sheet.Cells[firstRow + 24, 1]].
                Value2 = "Взам. инв. №";
            sheet.Range[sheet.Cells[firstRow + 25, 1], sheet.Cells[firstRow + 28, 1]].
                Value2 = "Подп. и дата";
            sheet.Range[sheet.Cells[firstRow + 29, 1], sheet.Cells[firstRow + 33, 1]].
                Value2 = "Инв. № подл.";

            // blank default text
            ((Excel.Range)sheet.Cells[firstRow + 33, 3]).Value2 = "Изм.";
            sheet.Range[sheet.Cells[firstRow + 33, 4], sheet.Cells[firstRow + 33, 5]].
                Value2 = "Лист";
            sheet.Range[sheet.Cells[firstRow + 33, 6], sheet.Cells[firstRow + 33, 8]].
                Value2 = "№ докум.";
            sheet.Range[sheet.Cells[firstRow + 33, 9], sheet.Cells[firstRow + 33, 10]].
                Value2 = "Подп.";
            sheet.Range[sheet.Cells[firstRow + 33, 11], sheet.Cells[firstRow + 33, 12]].
                Value2 = "Дата";
            sheet.Range[sheet.Cells[firstRow + 30, 26], sheet.Cells[firstRow + 31, 26]].
                Value2 = "Лист";
            sheet.Range[sheet.Cells[firstRow + 34, 13], sheet.Cells[firstRow + 34, 18]].
                Value2 = "Копировал:";
            sheet.Range[sheet.Cells[firstRow + 34, 19], sheet.Cells[firstRow + 34, 26]].
                Value2 = "Формат А4";
            sheet.Range[sheet.Cells[firstRow + 32, 26], sheet.Cells[firstRow + 33, 26]].
                Value2 = pageNumber;

            // font sizes
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 33, 2]].
                Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 12]].
                Font.Size = 11;
            ((Excel.Range)sheet.Cells[firstRow + 33, 3]).Font.Size = 10;
            sheet.Range[sheet.Cells[firstRow + 30, 11], sheet.Cells[firstRow + 32, 12]].
                Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow + 34, 1], sheet.Cells[firstRow + 34, 26]].
                Font.Size = 11;
            // обозначение
            sheet.Range[sheet.Cells[firstRow + 30, 13], sheet.Cells[firstRow + 33, 25]].
                Font.Size = 20;
            // лист
            sheet.Range[sheet.Cells[firstRow + 30, 26], sheet.Cells[firstRow + 31, 26]].
                Font.Size = 11;
            // номер листа
            sheet.Range[sheet.Cells[firstRow + 32, 26], sheet.Cells[firstRow + 33, 26]].
                Font.Size = 12;

            // text align
            //HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 33, 2]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 26]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 30, 13], sheet.Cells[firstRow + 33, 26]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 34, 1], sheet.Cells[firstRow + 34, 12]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 34, 13], sheet.Cells[firstRow + 34, 18]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range[sheet.Cells[firstRow + 34, 19], sheet.Cells[firstRow + 34, 26]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

        }

        virtual protected void FillBlankText()
        {
            A4SecondPageBlankTextFiller filler = new A4SecondPageBlankTextFiller();
            filler.Fill(sheet, firstRow);
        }
    }
}
