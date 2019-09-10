using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace DocGen.View.Unformatters
{
    class SpecUnformatter : Unformatter
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
                    if (cell.MergeArea.Count != 9 &&
                        cell.MergeArea.Count != 15)
                    {
                        delete = true;
                    }
                }
                else
                {
                    delete = true;
                }

                if ((int)cell.RowHeight >= 35)
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
                    } else
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

            string str = "Y1:Z1" + sep + "W1" + sep + "Q1:U1" +
                sep + "H1:O1" + sep + "F1" + sep + "A1:B1";
            range = sheet.Range[str];
            columns = range.EntireColumn;
            columns.Delete();

            // this code slow
            //Excel.Range cols = sheet.Columns;
            //int[] usedColumns = { 3, 4, 5, 7, 16, 22, 24 };
            //for (int i = 26; i >= 1; i--)
            //{
            //    if (!usedColumns.Contains(i))
            //    {
            //        range = (Excel.Range)cols[i];
            //        range.Delete();
            //    }
            //}
        }
    }
}
