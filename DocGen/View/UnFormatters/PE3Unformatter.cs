using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Unformatters
{
    class PE3Unformatter : Unformatter
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

            for (int i = usedRange.Rows.Count; i >= 1; i--)
            {
                cell = (Excel.Range)usedRange.Cells[i, 8];
                if ((bool)cell.MergeCells)
                {
                    if (cell.MergeArea.Count != 13)
                    {
                        delete = true;
                    }
                } else
                {
                    delete = true;
                }

                if ((int)cell.RowHeight >= 35)
                {
                    delete = true;
                }

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

            string str = "X1:Z1" + sep + "V1" + sep + "I1:T1" +
                sep + "E1:G1" + sep + "A1:B1";
            range = sheet.Range[str];
            columns = range.EntireColumn;
            columns.Delete();

            // this code slow
            // used columns
            //int[] columns = { 3, 4, 8, 21, 23 };
            //for (int i = 26; i >= 1; i--)
            //{
            //    if (!columns.Contains(i))
            //    {
            //        ((Excel.Range)sheet.Columns[i]).Delete();
            //    }
            //}
        }
    }
}
