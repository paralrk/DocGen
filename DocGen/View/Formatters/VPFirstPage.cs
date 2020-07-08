using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class VPFirstPage : A3FirstPage
    {
        public VPFirstPage(Excel.Worksheet sheet, int pageCount) : base(sheet, pageCount)
        {
        }

        override protected void SetColumnsWidth()
        {
            // insert extra columns
            Excel.Range range;


            range = sheet.Range["A1:B1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["E1:N1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["P1:Q1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["S1:Z1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["AB1:AD1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["AF1:AJ1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["AN1:AO1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["AQ1:AQ1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["AS1:AT1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);


            base.SetColumnsWidth();
        }
        override protected void MergeCells()
        {
            sheet.Range["A1:AT2"].UnMerge();

            base.MergeCells();

            // specific for VP
            //title
            sheet.Range["C1:C2"].Merge();
            sheet.Range["D1:N2"].Merge();
            sheet.Range["O1:Q2"].Merge();
            sheet.Range["R1:Z2"].Merge();
            sheet.Range["AA1:AD2"].Merge();
            sheet.Range["AE1:AJ2"].Merge();
            sheet.Range["AK1:AQ1"].Merge(); // Количество
            sheet.Range["AM2:AO2"].Merge();
            sheet.Range["AP2:AQ2"].Merge();
            sheet.Range["AR1:AT2"].Merge();
            //rows
            for (int i = 3; i <= 25; i++)
            {
                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 14]].Merge();
                sheet.Range[sheet.Cells[i, 15], sheet.Cells[i, 17]].Merge();
                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 26]].Merge();
                sheet.Range[sheet.Cells[i, 27], sheet.Cells[i, 30]].Merge();
                sheet.Range[sheet.Cells[i, 31], sheet.Cells[i, 36]].Merge();
                sheet.Range[sheet.Cells[i, 39], sheet.Cells[i, 41]].Merge();
                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 46]].Merge();
            }
        }

        override protected void DrawBorders()
        {
            base.DrawBorders();

            // specific for VP
            // title
            sheet.Range["C1:AT2"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range["C3:AT25"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["C3:AT25"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные            
        }

        override protected void FillBlank()
        {
            base.FillBlank();

            // specific for VP
            sheet.Range["C1:C2"].Orientation = 90;
            // title default text
            sheet.Range["C1:C2"].Value2 = "№ строки";
            sheet.Range["D1:N2"].Value2 = "Наименование";
            sheet.Range["O1:Q2"].Value2 = "Код продукции";
            sheet.Range["R1:Z2"].Value2 = "Обозначение документа";
            sheet.Range["AA1:AD2"].Value2 = "Поставщик";
            sheet.Range["AE1:AJ2"].Value2 = "Куда входит (обозначение)";
            sheet.Range["AK1:AQ1"].Value2 = "Количество";
            sheet.Range["AK2:AK2"].Value2 = "на из- делие";
            sheet.Range["AL2:AL2"].Value2 = "в ком- плекты";
            sheet.Range["AM2:AO2"].Value2 = "на ре- гулир.";
            sheet.Range["AP2:AQ2"].Value2 = "всего";
            sheet.Range["AR1:AT2"].Value2 = "Приме- чание";
            sheet.Range["AH36:AM36"].Value2 = "Ведомость покупных изделий";

            for (int i = 3, line = 1; i <= 25; i++, line++)
            {
                ((Excel.Range)sheet.Cells[i, 3]).Value2 = line;
            }

            // font sizes
            sheet.Range["C1:C2"].Font.Size = 11;
            sheet.Range["D1:AT2"].Font.Size = 14;
            sheet.Range["AK2:AQ2"].Font.Size = 11;
            sheet.Range["D1:AT2"].WrapText = true;

            // text align
            sheet.Range["C1:AT2"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["C1:AT2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            sheet.Range["C3:AT25"].ShrinkToFit = true;
        }
    }
}
