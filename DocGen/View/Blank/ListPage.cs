using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

using DocGen.Utils;

namespace DocGen.View.Blank
{
    static class ListPage
    {
        public static Excel.Worksheet GetListPage()
        {
            Excel.Worksheet sheet;
            string name = "Содержание";
            if (!SheetHelper.isExisting(name))
            {
                sheet = SheetHelper.getSheet(name);
                FillDefault(sheet);
                FormatCells(sheet);
            }
            else
            {
                sheet = SheetHelper.getSheet(name);
            }
            return sheet;
        }

        public static void FillDefault(Excel.Worksheet sheet)
        {
            Dictionary<int, string> namesList = new Dictionary<int, string> {
                { (int)Names.DocumentType, "Тип документа" },
                { (int)Names.Sheet, "Лист" },
                { (int)Names.Designation, "Обозначение" },
                { (int)Names.Name, "Наименование" },
                { (int)Names.Product, "Изделие" },
                { (int)Names.PrimaryUse, "Первичное применение" },
                { (int)Names.Developed, "Разработал" },
                { (int)Names.Cheked, "Проверил" },
                { (int)Names.AdditionalField, "ДопПоле" },
                { (int)Names.AdditionalSurname, "ДопФамилия" },
                { (int)Names.NormalControl, "Нконтр" },
                { (int)Names.Approved, "Утв" },
                { (int)Names.DateDeveloped, "ДатаРазработал" },
                { (int)Names.DateCheked, "ДатаПроверил" },
                { (int)Names.DateAdditionalSurname, "ДатаДопФамили" },
                { (int)Names.DateNormalControl, "ДатаНконтр" },
                { (int)Names.DateApproved, "ДатаУтв" },
                { (int)Names.Letter1, "Литера1" },
                { (int)Names.Letter2, "Литера2" },
                { (int)Names.Letter3, "Литера3" },
                { (int)Names.RefNumber, "Справ№" },
                { (int)Names.OriginalInvNumber, "Инв№Подл" },
                { (int)Names.SignDateOriginal, "ПодпИДатаПодл" },
                { (int)Names.InsteadInvNumber, "ВзамИнв№" },
                { (int)Names.DublicateInvNumber, "Инв№Дубл" },
                { (int)Names.SignDateDublicate, "ПодпИДатаДубл" },
                { (int)Names.DesignationLU, "ОбозначениеЛУ" },
                { (int)Names.ApprovalSheet, "УтвЛит" },
                { (int)Names.ApprovalDoc, "УтвДок" },
                { (int)Names.CustomerIndex, "ИндЗаказчика" },
                { (int)Names.Firm, "Фирма" }
            };

            for (int i = 1; i <= namesList.Count; i++)
            {
                Excel.Range cells = (Excel.Range)sheet.Cells[i, 1];
                cells.Value2 = namesList[i];
            }
        }

        public static void FormatCells(Excel.Worksheet sheet)
        {
            Excel.Style style = StyleHelper.getDocGenMainStyle();
            sheet.Cells.Style = style;
            sheet.Cells.ColumnWidth = 30;
            sheet.Cells.Borders[Excel.XlBordersIndex.xlInsideVertical].
                Weight = Excel.XlBorderWeight.xlThin;
            ((Excel.Range)sheet.Rows[(int)Names.Name]).RowHeight = 60;
            ((Excel.Range)sheet.Rows[(int)Names.Name]).WrapText = true;
        }

        public static void AddDocument(string name, string documentType)
        {
            Excel.Worksheet sheet = GetListPage();
            Excel.Range usedColumns = (Excel.Range)sheet.UsedRange.Rows;
            int column = usedColumns.Columns.Count + 1;

            // sheet name
            Excel.Range cells = (Excel.Range)sheet.Cells[(int)Names.Sheet, column];
            cells.Value2 = name;

            // document's type
            cells = (Excel.Range)sheet.Cells[(int)Names.DocumentType, column];
            cells.Value2 = documentType;

        }

        public static string GetDocumentColumn(string name)
        {
            Excel.Worksheet sheet = GetListPage();
            Excel.Range nameRow = (Excel.Range)sheet.Rows[2];
            Excel.Range columns;
            int usedColumns = sheet.UsedRange.Columns.Count;
            string column = null;
            Regex regex = new Regex(@"[a-zA-Z]+");
            Match match = null;
            for (int i = 2; i <= usedColumns; i++)
            {
                columns = (Excel.Range)nameRow.Columns[i];
                if (name.Equals(columns.Value2))
                {
                    match = regex.Match(columns.Address);
                    column = match.Groups[0].Value;
                }
            }
            return column;
        }

        public static int GetDocumentColumnNumber(string name)
        {
            Excel.Worksheet sheet = GetListPage();
            Excel.Range nameRow = (Excel.Range)sheet.Rows[2];
            Excel.Range columns;
            int usedColumns = sheet.UsedRange.Columns.Count;
            int column = -1;
            for (int i = 2; i <= usedColumns; i++)
            {
                columns = (Excel.Range)nameRow.Columns[i];
                if (name.Equals(columns.Value2))
                {
                    column = i;
                }
            }
            return column;
        }

        public static string GetDocumentType (Excel.Worksheet sheet)
        {
            string type = null;
            int column = GetDocumentColumnNumber(sheet.Name);
            if (column > 1)
            {
                Excel.Worksheet listPage = GetListPage();
                Excel.Range cells = (Excel.Range)listPage.Cells[(int)Names.DocumentType, column];
                type = (string)cells.Value2;
            }
            return type;
        }

    }
}
