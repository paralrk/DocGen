using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace DocGen.View.Unformatters
{
    class D33_UDUnformatter : Unformatter
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

                cell = (Excel.Range)usedRange.Cells[i, 7];
                if ((bool)cell.MergeCells)
                {
                    if (cell.MergeArea.Count != 7)
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

            string str = "V1:Z1" + sep + "R1:T1" + sep + "O1:P1" +
                sep + "K1:M1" + sep + "D1:I1" + sep + "A1:B1";
            range = sheet.Range[str];
            columns = range.EntireColumn;
            columns.Delete();

        }
    }
}
