using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using DocGen.Utils;

namespace DocGen.View.Blank
{
    class A4FirstPageBlankTextFiller
    {
        public void Fill(Excel.Worksheet sheet)
        {
            Excel.Worksheet list = ListPage.GetListPage();
            string column = ListPage.GetDocumentColumn(sheet.Name);
            // for example: Value2 = "=Содержание!$B$3"
            sheet.Range["M29:Z31"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Designation; // Обозначение
            sheet.Range["M32:R35"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Name; // Наименование
            sheet.Range["A14:B15"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Product; // Изделие
            sheet.Range["B1:B6"].Value2   = "=Содержание!$" + column + "$" + (int)Names.PrimaryUse; // Первичное применение
            sheet.Range["F32:H32"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Developed; // Разработал
            sheet.Range["F33:H33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Cheked; // Проверил
            sheet.Range["C34:E34"].Value2 = "=Содержание!$" + column + "$" + (int)Names.AdditionalField; // ДопПоле
            sheet.Range["F34:H34"].Value2 = "=Содержание!$" + column + "$" + (int)Names.AdditionalSurname; // ДопФамилия
            sheet.Range["F35:H35"].Value2 = "=Содержание!$" + column + "$" + (int)Names.NormalControl; // Нконтр
            sheet.Range["F36:H36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Approved; // Утв
            sheet.Range["K32:L32"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateDeveloped; // ДатаРазработал
            sheet.Range["K33:L33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateCheked; // ДатаПроверил
            sheet.Range["K34:L34"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateAdditionalSurname; // ДатаДопФамилия
            sheet.Range["K35:L35"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateNormalControl; // ДатаНконтр
            sheet.Range["K36:L36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateApproved; // ДатаУтв
            sheet.Range["S33:S33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Letter1; // Литера1
            sheet.Range["T33:T33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Letter2; // Литера2
            sheet.Range["U33:U33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Letter3; // Литера3
            sheet.Range["B7:B13"].Value2  = "=Содержание!$" + column + "$" + (int)Names.RefNumber; // Справ№
            sheet.Range["B32:B36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.OriginalInvNumber; // Инв№Подл
            sheet.Range["B27:B31"].Value2 = "=Содержание!$" + column + "$" + (int)Names.SignDateOriginal; // ПодпИДатаПодл
            sheet.Range["B23:B26"].Value2 = "=Содержание!$" + column + "$" + (int)Names.InsteadInvNumber; // ВзамИнв№
            sheet.Range["B20:B22"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DublicateInvNumber; // Инв№Дубл
            sheet.Range["B16:B19"].Value2 = "=Содержание!$" + column + "$" + (int)Names.SignDateDublicate; // ПодпИДатаДубл
            sheet.Range["C26:L28"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DesignationLU; // ОбозначениеЛУ
            sheet.Range["O26:R27"].Value2 = "=Содержание!$" + column + "$" + (int)Names.ApprovalSheet; // УтвЛист
            sheet.Range["S26:Z27"].Value2 = "=Содержание!$" + column + "$" + (int)Names.ApprovalDoc; // УтвДок
            sheet.Range["M28:Z28"].Value2 = "=Содержание!$" + column + "$" + (int)Names.CustomerIndex; // ИндЗаказчика
            sheet.Range["S34:Z36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Firm; // Фирма

        }
    }
}
