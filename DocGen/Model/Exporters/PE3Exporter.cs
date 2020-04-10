using DocGen.View.EmptyDocuments;
using DocGen.Model.Documents;
using DocGen.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.Model.Exporters
{
    class PE3Exporter : IExporter
    {
        public void Export(IDocument pe3)
        {
            ExcelHelper.DisableUpdating();

            EmptyDocument pe3Empty = new PE3EmptyDocument();
            pe3Empty.NewDocument();
            pe3Empty.InitFormatCells();
            pe3Empty.Format();
            Excel.Worksheet sheet = pe3Empty.getSheet();

            int rowNumber = 3;
            List<IRow> rows = pe3.CombineRows();

            foreach (RowPE3 row in rows)
            {
                sheet.Cells[rowNumber, 1] = row.Zone;
                sheet.Cells[rowNumber, 2] = row.Designator;
                sheet.Cells[rowNumber, 3] = row.Name;
                if (row.Quantity != 0)
                {
                    sheet.Cells[rowNumber, 4] = row.Quantity;
                }
                sheet.Cells[rowNumber, 5] = row.Note;
                rowNumber++;
            }

            ExcelHelper.EnableUpdating();
        }
    }
}
