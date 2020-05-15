using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace DocGen.View.Unformatters
{
    class VPUnformatter : Unformatter
    {
        protected override void DeleteRows()
        {
            base.DeleteRows();

            Excel.Range usedRange = sheet.UsedRange;
            Excel.Range cell = null;
            Excel.Range range;
            bool delete = false;

            // separator can be ; or ,
            var sep = (string)Globals.ThisAddIn.Application.
                International[Excel.XlApplicationInternational.xlListSeparator];
            string str = "";

            Stopwatch sw = new Stopwatch();
            Debug.WriteLine("--------------- Test performance -------------");
            for (int i = usedRange.Rows.Count; i >= 1; i--)
            {

                cell = (Excel.Range)usedRange.Cells[i, 7];
                if ((bool)cell.MergeCells)
                {
                    if (cell.MergeArea.Count != 11 &&
                        cell.MergeArea.Count != 33 &&
                        cell.MergeArea.Count != 40)
                    {
                        delete = true;
                    }
                }
                else
                {
                    delete = true;
                }

                if ((int)cell.RowHeight >= 35 || (int)cell.RowHeight < 25)
                {
                    delete = true;
                }

                sw.Start();
                if (delete)
                {
                    str += i + ":" + i;
                    if (str.Length < 200)
                    {
                        str += sep;
                    }
                    else
                    {
                        range = sheet.Range[str];
                        range.Delete();
                        str = "";
                    }
                    delete = false;
                }
                sw.Stop();
                Debug.WriteLine("Check Delete() Elapsed={0}", sw.Elapsed);
                sw.Reset();
            }

            if (str.Length > 1)
            {
                // delete separator at the end of str
                str = str.Remove(str.Length - 1);
                range = sheet.Range[str];
                range.Delete();
            }
        }

        protected override void DeleteColumns()
        {
            base.DeleteColumns();

            // separator can be ; or ,
            var sep = (string)Globals.ThisAddIn.Application.
                International[Excel.XlApplicationInternational.xlListSeparator];

            Excel.Range range;
            Excel.Range columns;

            string str = "AS1:AT1" + sep + "AQ1" + sep + "AN1:AO1"
                + sep + "AF1:AJ1" + sep + "AB1:AD1" + sep + "S1:Z1"
                + sep + "P1:Q1" + sep + "E1:N1" + sep + "A1:B1";
            range = sheet.Range[str];
            columns = range.EntireColumn;
            columns.Delete();

            sheet.Range["A1:A1"].EntireColumn.ClearContents();

        }
    }
}
