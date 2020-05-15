using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.View.Blank;
using System.Diagnostics;

namespace DocGen.View
{
    class SheetBordersDrawer
    {
        private int firstPageRows = 23;
        private int secondPageRows = 29;
        private int columnsCount = 5;
        Excel.Worksheet sheet;
        private int previousRowsCount = 0;

        public SheetBordersDrawer()
        {
            //Initialize();
        }
        private void Initialize()
        {
            sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string type = ListPage.GetDocumentType(sheet);
            switch (type)
            {
                case "Перечень элементов":
                    firstPageRows = 23;
                    secondPageRows = 29;
                    columnsCount = 5;
                    break;
                case "Спецификация":
                    firstPageRows = 23;
                    secondPageRows = 29;
                    columnsCount = 7;
                    break;
                case "Ведомость покупных изделий":
                    firstPageRows = 24;
                    secondPageRows = 28;
                    columnsCount = 11;
                    break;
                default:
                    break;
            }
        }

        public void EnableSheetChangeEvent()
        {
            Initialize();
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            sheet.Change += Sheet_Change;
        }

        private void Sheet_Change(Excel.Range Target)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range usedRange = sheet.UsedRange;
            int usedRows = usedRange.Rows.Count;
            if (previousRowsCount != usedRows)
            {
                DrawSheetBorders();
                previousRowsCount = usedRows;
            }
        }

        public void DisableSheetChangeEvent()
        {
            Initialize();
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            sheet.Change -= Sheet_Change;
            // DeleteSheetBorders();
        }

        public void DrawSheetBorders()
        {
            sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range usedRange = (Excel.Range)sheet.UsedRange;
            int usedRows = usedRange.Rows.Count;
            //int usedColumns = usedRange.Columns.Count;
            int drawnRow = 0;
            Excel.Range row;

            DeleteSheetBorders();

            if (usedRows > firstPageRows)
            {
                // 1 - row with header
                drawnRow = firstPageRows + 1;
                row = sheet.Range[sheet.Cells[drawnRow, 1],
                    sheet.Cells[drawnRow, columnsCount]];
                //row = (Excel.Range)sheet.Rows[drawnRow];
                row.Borders[Excel.XlBordersIndex.xlEdgeBottom].
                    LineStyle = Excel.XlLineStyle.xlDashDot;
                row.Borders[Excel.XlBordersIndex.xlEdgeBottom].
                    Weight = Excel.XlBorderWeight.xlMedium;
            }
            while (drawnRow + secondPageRows < usedRows)
            {
                drawnRow += secondPageRows;
                row = sheet.Range[sheet.Cells[drawnRow, 1],
                    sheet.Cells[drawnRow, columnsCount]];
                //row = (Excel.Range)sheet.Rows[drawnRow];
                row.Borders[Excel.XlBordersIndex.xlEdgeBottom].
                    LineStyle = Excel.XlLineStyle.xlDashDot;
                row.Borders[Excel.XlBordersIndex.xlEdgeBottom].
                    Weight = Excel.XlBorderWeight.xlMedium;
            }
            Excel.Range column = sheet.Range[sheet.Cells[1, columnsCount],
                    sheet.Cells[usedRows, columnsCount]];
            column.Borders[Excel.XlBordersIndex.xlEdgeRight].
                    LineStyle = Excel.XlLineStyle.xlDashDot;
            column.Borders[Excel.XlBordersIndex.xlEdgeRight].
                Weight = Excel.XlBorderWeight.xlMedium;
        }



        public void DeleteSheetBorders()
        {
            sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range usedRange = (Excel.Range)sheet.UsedRange;
            if (usedRange != null)
            {
                usedRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            }
        }
    }
}
