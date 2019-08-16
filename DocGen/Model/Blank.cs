using DocGen.Utils;
using DocGen.View.Blank;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.Model
{
    class Blank
    {
        public Dictionary<int, string> NamesList { get; set; }

        private Excel.Worksheet sheet;

        public Blank(Excel.Worksheet sheet)
        {
            this.sheet = sheet;

            NamesList = new Dictionary<int, string> {
                { (int)Names.DocumentType, "" },
                { (int)Names.Sheet, "" },
                { (int)Names.Designation, "" },
                { (int)Names.Name, "" },
                { (int)Names.Product, "" },
                { (int)Names.PrimaryUse, "" },
                { (int)Names.Developed, "" },
                { (int)Names.Cheked, "" },
                { (int)Names.AdditionalField, "" },
                { (int)Names.AdditionalSurname, "" },
                { (int)Names.NormalControl, "" },
                { (int)Names.Approved, "" },
                { (int)Names.DateDeveloped, "" },
                { (int)Names.DateCheked, "" },
                { (int)Names.DateAdditionalSurname, "" },
                { (int)Names.DateNormalControl, "" },
                { (int)Names.DateApproved, "" },
                { (int)Names.Letter1, "" },
                { (int)Names.Letter2, "" },
                { (int)Names.Letter3, "" },
                { (int)Names.RefNumber, "" },
                { (int)Names.OriginalInvNumber, "" },
                { (int)Names.SignDateOriginal, "" },
                { (int)Names.InsteadInvNumber, "" },
                { (int)Names.DublicateInvNumber, "" },
                { (int)Names.SignDateDublicate, "" },
                { (int)Names.DesignationLU, "" },
                { (int)Names.ApprovalSheet, "" },
                { (int)Names.ApprovalDoc, "" },
                { (int)Names.CustomerIndex, "" },
                { (int)Names.Firm, "" }
            };
        }

        public void FillListPage()
        {
            Excel.Worksheet listSheet = ListPage.GetListPage();
            int column = ListPage.GetDocumentColumnNumber(sheet.Name);

            if (column > 1)
            {
                foreach (KeyValuePair<int, string> entry in NamesList)
                {
                    Excel.Range cells = (Excel.Range)listSheet.Cells[entry.Key, column];
                    cells.Value2 = entry.Value;
                }

                string name = NamesList[(int)Names.Designation];
                if (!String.IsNullOrEmpty(name))
                {
                    NamesList[(int)Names.Sheet] = name;
                    sheet.Name = name;
                }
                else
                {
                    Excel.Range cells = (Excel.Range)listSheet.Cells[(int)Names.Sheet, column];
                    cells.Value2 = sheet.Name;
                }

            }
        }

        public void ReadListPage()
        {
            Excel.Worksheet listSheet = ListPage.GetListPage();
            int column = ListPage.GetDocumentColumnNumber(sheet.Name);

            if (column > 1)
            {
                List<int> ids = NamesList.Keys.ToList();
                foreach (int idKey in ids)
                {
                    Excel.Range cells = (Excel.Range)listSheet.Cells[idKey, column];
                    NamesList[idKey] = (string)cells.Value2;
                }

            }

        }

    }
}
