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
    class SpecificationExporter : IExporter
    {
        public void Export(Document spec)
        {
            ExcelHelper.DisableUpdating();

            EmptyDocumentsFactory factory = new EmptyDocumentsFactory();
            EmptyDocument specEmpty = factory.GetSpecificationEmptyDocument();
            specEmpty.NewDocument();
            specEmpty.Format();
            Excel.Worksheet sheet = specEmpty.getSheet();

            int rowNumber = 3;
            List<IRow> rows = spec.CombineRows();

            foreach (RowSpec row in rows)
            {
                sheet.Cells[rowNumber, 1] = row.Format;
                sheet.Cells[rowNumber, 2] = row.Zone;
                sheet.Cells[rowNumber, 3] = row.Position;
                sheet.Cells[rowNumber, 4] = row.Designation;
                sheet.Cells[rowNumber, 5] = row.Name;
                if (row.Quantity != 0)
                {
                    sheet.Cells[rowNumber, 6] = row.Quantity;
                }
                sheet.Cells[rowNumber, 7] = row.Note;
                rowNumber++;
            }
            ExcelHelper.EnableUpdating();
        }
    }
}
