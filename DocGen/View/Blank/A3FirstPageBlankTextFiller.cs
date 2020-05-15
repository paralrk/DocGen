using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using DocGen.Utils;

namespace DocGen.View.Blank
{
    class A3FirstPageBlankTextFiller
    {
        public void Fill(Excel.Worksheet sheet)
        {
            Excel.Worksheet list = ListPage.GetListPage();
            string column = ListPage.GetDocumentColumn(sheet.Name);
            // for example: Value2 = "=Содержание!$B$3"
            sheet.Range["AH29:AT31"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Designation; // Обозначение
            sheet.Range["AH32:AM35"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Name; // Наименование
            sheet.Range["A14:B15"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Product; // Изделие
            sheet.Range["B1:B6"].Value2 = "=Содержание!$" + column + "$" + (int)Names.PrimaryUse; // Первичное применение
            sheet.Range["AD32:AE32"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Developed; // Разработал
            sheet.Range["AD33:AE33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Cheked; // Проверил
            sheet.Range["AB34:AC34"].Value2 = "=Содержание!$" + column + "$" + (int)Names.AdditionalField; // ДопПоле
            sheet.Range["AD34:AE34"].Value2 = "=Содержание!$" + column + "$" + (int)Names.AdditionalSurname; // ДопФамилия
            sheet.Range["AD35:AE35"].Value2 = "=Содержание!$" + column + "$" + (int)Names.NormalControl; // Нконтр
            sheet.Range["AD36:AE36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Approved; // Утв
            sheet.Range["AG32:AG32"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateDeveloped; // ДатаРазработал
            sheet.Range["AG33:AG33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateCheked; // ДатаПроверил
            sheet.Range["AG34:AG34"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateAdditionalSurname; // ДатаДопФамилия
            sheet.Range["AG35:AG35"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateNormalControl; // ДатаНконтр
            sheet.Range["AG36:AG36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DateApproved; // ДатаУтв
            sheet.Range["AN33:AN33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Letter1; // Литера1
            sheet.Range["AO33:AO33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Letter2; // Литера2
            sheet.Range["AP33:AP33"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Letter3; // Литера3
            sheet.Range["B7:B13"].Value2 = "=Содержание!$" + column + "$" + (int)Names.RefNumber; // Справ№
            sheet.Range["B32:B36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.OriginalInvNumber; // Инв№Подл
            sheet.Range["B27:B31"].Value2 = "=Содержание!$" + column + "$" + (int)Names.SignDateOriginal; // ПодпИДатаПодл
            sheet.Range["B23:B26"].Value2 = "=Содержание!$" + column + "$" + (int)Names.InsteadInvNumber; // ВзамИнв№
            sheet.Range["B20:B22"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DublicateInvNumber; // Инв№Дубл
            sheet.Range["B16:B19"].Value2 = "=Содержание!$" + column + "$" + (int)Names.SignDateDublicate; // ПодпИДатаДубл
            sheet.Range["AB26:AG28"].Value2 = "=Содержание!$" + column + "$" + (int)Names.DesignationLU; // ОбозначениеЛУ
            sheet.Range["AI26:AL27"].Value2 = "=Содержание!$" + column + "$" + (int)Names.ApprovalSheet; // УтвЛист
            sheet.Range["AM26:AT27"].Value2 = "=Содержание!$" + column + "$" + (int)Names.ApprovalDoc; // УтвДок
            sheet.Range["AH28:AT28"].Value2 = "=Содержание!$" + column + "$" + (int)Names.CustomerIndex; // ИндЗаказчика
            sheet.Range["AN34:AT36"].Value2 = "=Содержание!$" + column + "$" + (int)Names.Firm; // Фирма
        }
    }
}
