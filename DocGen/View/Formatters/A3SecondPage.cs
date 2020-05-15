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
    abstract class A3SecondPage
    {
        protected Excel.Worksheet sheet;
        public int Height { get; } = 35;
        public int RowsCount { get; } = 28;
        protected int pageNumber = 2;
        protected int firstRow;

        public A3SecondPage(Excel.Worksheet sheet, int firstRow, int pageNumber)
        {
            this.sheet = sheet;
            this.firstRow = firstRow;
            this.pageNumber = pageNumber;
        }

        public void Format()
        {
            Stopwatch sw = new Stopwatch();
            Debug.WriteLine("Formatting VP Second Page document");

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
            string rowsRange = firstRow + ":" + (firstRow + 1) + sep +
                    (firstRow + 29) + ":" + (firstRow + 33);
            sheet.Range[rowsRange].Insert();

            string str;
            Excel.Range range;
            str = firstRow + ":" + firstRow;
            range = sheet.Range[str];
            range.RowHeight = 24;
            str = (firstRow + 1) + ":" + (firstRow + 1);
            range = sheet.Range[str];
            range.RowHeight = 50;
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
            // common for A3 blank
            // left cells
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 14, 2]].Merge();
            // empty cells
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 27]].Merge();
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
            sheet.Range[sheet.Cells[firstRow + 34, 1], sheet.Cells[firstRow + 34, 33]].Merge();
            sheet.Range[sheet.Cells[firstRow + 34, 34], sheet.Cells[firstRow + 34, 39]].Merge();
            sheet.Range[sheet.Cells[firstRow + 34, 40], sheet.Cells[firstRow + 34, 46]].Merge();
            //blank

            // Изм.
            sheet.Range[sheet.Cells[firstRow + 31, 28], sheet.Cells[firstRow + 32, 28]].Merge();

            // Лист 
            sheet.Range[sheet.Cells[firstRow + 31, 29], sheet.Cells[firstRow + 32, 29]].Merge();

            // № докум
            sheet.Range[sheet.Cells[firstRow + 30, 30], sheet.Cells[firstRow + 30, 31]].Merge();
            sheet.Range[sheet.Cells[firstRow + 31, 30], sheet.Cells[firstRow + 32, 31]].Merge();
            sheet.Range[sheet.Cells[firstRow + 33, 30], sheet.Cells[firstRow + 33, 31]].Merge();

            // Подп.
            sheet.Range[sheet.Cells[firstRow + 31, 32], sheet.Cells[firstRow + 32, 32]].Merge();

            // Дата
            sheet.Range[sheet.Cells[firstRow + 31, 33], sheet.Cells[firstRow + 32, 33]].Merge();

            // обозначение
            sheet.Range[sheet.Cells[firstRow + 30, 34], sheet.Cells[firstRow + 33, 45]].Merge();
            // лист
            sheet.Range[sheet.Cells[firstRow + 30, 46], sheet.Cells[firstRow + 31, 46]].Merge();
            // номер листа
            sheet.Range[sheet.Cells[firstRow + 32, 46], sheet.Cells[firstRow + 33, 46]].Merge();
        }

        virtual protected void DrawBorders()
        {
            // common for A3 blank second page
            // clear borders
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 34, 46]].
                Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //left cells
            sheet.Range[sheet.Cells[firstRow + 15, 1], sheet.Cells[firstRow + 33, 2]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;

            // blank
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 46]].
                            Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range[sheet.Cells[firstRow + 30, 27], sheet.Cells[firstRow + 33, 33]].
                            Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;  // внутренние горизонтальные

        }

        virtual protected void FillBlank()
        {
            // common for A3 blank second page
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
            ((Excel.Range)sheet.Cells[firstRow + 33, 28]).Value2 = "Изм.";
            ((Excel.Range)sheet.Cells[firstRow + 33, 29]).Value2 = "Лист";
            sheet.Range[sheet.Cells[firstRow + 33, 30], sheet.Cells[firstRow + 33, 31]].
                Value2 = "№ докум.";
            ((Excel.Range)sheet.Cells[firstRow + 33, 32]).Value2 = "Подп.";
            ((Excel.Range)sheet.Cells[firstRow + 33, 32]).Value2 = "Дата";
            sheet.Range[sheet.Cells[firstRow + 30, 46], sheet.Cells[firstRow + 31, 46]].
                Value2 = "Лист";
            sheet.Range[sheet.Cells[firstRow + 34, 34], sheet.Cells[firstRow + 34, 39]].
                Value2 = "Копировал:";
            sheet.Range[sheet.Cells[firstRow + 34, 40], sheet.Cells[firstRow + 34, 46]].
                Value2 = "Формат А4";
            sheet.Range[sheet.Cells[firstRow + 32, 46], sheet.Cells[firstRow + 33, 46]].
                Value2 = pageNumber;

            // font sizes
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 33, 2]].
                Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow + 30, 27], sheet.Cells[firstRow + 33, 32]].
                Font.Size = 11;
            ((Excel.Range)sheet.Cells[firstRow + 33, 28]).Font.Size = 10; // изм.
            sheet.Range[sheet.Cells[firstRow + 34, 1], sheet.Cells[firstRow + 34, 46]].
                Font.Size = 11;
            // обозначение
            sheet.Range[sheet.Cells[firstRow + 30, 34], sheet.Cells[firstRow + 33, 45]].
                Font.Size = 20;
            // лист
            sheet.Range[sheet.Cells[firstRow + 30, 46], sheet.Cells[firstRow + 31, 46]].
                Font.Size = 11;
            // номер листа
            sheet.Range[sheet.Cells[firstRow + 32, 46], sheet.Cells[firstRow + 33, 46]].
                Font.Size = 12;

            // text align
            sheet.Range[sheet.Cells[firstRow, 1], sheet.Cells[firstRow + 33, 2]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 30, 3], sheet.Cells[firstRow + 33, 46]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 30, 33], sheet.Cells[firstRow + 33, 46]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 34, 1], sheet.Cells[firstRow + 34, 33]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 34, 34], sheet.Cells[firstRow + 34, 39]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range[sheet.Cells[firstRow + 34, 40], sheet.Cells[firstRow + 34, 46]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

        }

        virtual protected void FillBlankText()
        {
            A3SecondPageBlankTextFiller filler = new A3SecondPageBlankTextFiller();
            filler.Fill(sheet, firstRow);
        }
    }
}
