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
    class VPExporter : IExporter
    {
        public void Export(IDocument vp)
        {
            ExcelHelper.DisableUpdating();

            EmptyDocumentsFactory factory = new EmptyDocumentsFactory();
            EmptyDocument vpEmpty = factory.GetVPEmptyDocument();
            vpEmpty.NewDocument();
            vpEmpty.InitFormatCells();
            vpEmpty.Format();
            Excel.Worksheet sheet = vpEmpty.getSheet();

            int rowNumber = 4;
            List<IRow> rows = vp.CombineRows();

            foreach (RowVP row in rows)
            {
                if (row.RowNumber != 0)
                {
                    sheet.Cells[rowNumber, 1] = row.RowNumber;
                }
                sheet.Cells[rowNumber, 2] = row.Name;
                sheet.Cells[rowNumber, 3] = row.ProductCode;
                sheet.Cells[rowNumber, 4] = row.Designation;
                sheet.Cells[rowNumber, 5] = row.Supplier;
                sheet.Cells[rowNumber, 6] = row.WhereItUsed;
                if (row.QuantityProduct != 0)
                {
                    sheet.Cells[rowNumber, 7] = row.QuantityProduct;
                }
                if (row.QuantitySet != 0)
                {
                    sheet.Cells[rowNumber, 8] = row.QuantitySet;
                }
                if (row.QuantityAdjustment != 0)
                {
                    sheet.Cells[rowNumber, 9] = row.QuantityAdjustment;
                }
                if (row.QuantityTotal != 0)
                {
                    sheet.Cells[rowNumber, 10] = row.QuantityTotal;
                }
                sheet.Cells[rowNumber, 11] = row.Note;
                rowNumber++;
            }
            ExcelHelper.EnableUpdating();
        }

    }
}
