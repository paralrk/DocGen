using DocGen.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.Utils
{
    class CellsSizeManager
    {
        public int ColumnsWidth { get; set; }

        public CellsSizeManager()
        {
             //ColumnsWidth = settings.ColumnsWidth;
        }

        public void SetColumnWidth(Excel.Worksheet sheet,  int widthScale)
        {
            int usedColumns = sheet.UsedRange.Columns.Count;
            double scale = widthScale / 100.0;
            Excel.Range column;
            double width;

            for (int i = 1; i <= usedColumns; i++)
            {
                column = (Excel.Range)sheet.Cells[1, i];
                width = (double)column.ColumnWidth;
                column.ColumnWidth = width * scale;
            }
        }

        public void SetRowsHeight(Excel.Worksheet sheet, int heightScale)
        {
            int usedRows = sheet.UsedRange.Rows.Count;
            double scale = heightScale / 100.0;
            Excel.Range row;
            double height;

            for (int i = 1; i <= usedRows; i++)
            {
                row = (Excel.Range)sheet.Cells[i, 1];
                height = (double)row.RowHeight;
                row.RowHeight = height * scale;
            }
        }
    }
}
