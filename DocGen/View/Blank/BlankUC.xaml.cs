using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.Utils;

namespace DocGen.View.Blank
{
    /// <summary>
    /// Interaction logic for BlankUC.xaml
    /// </summary>
    public partial class BlankUC : UserControl
    {
        private Model.Blank blank;

        public BlankUC()
        {
            InitializeComponent();
            //blank = new Blank((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            //FillTextBoxes();
        }


        public void FillTextBoxes()
        {
            blank = new Model.Blank((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            blank.ReadListPage();
            Dictionary<int, string> namesList = blank.NamesList;

            labelDocumentType.Content = namesList[(int)Names.DocumentType];
            //sheet.Text = blank.DocumentType;
            designation.Text = namesList[(int)Names.Designation];
            name.Text = namesList[(int)Names.Name];
            product.Text = namesList[(int)Names.Product];
            primaryUse.Text = namesList[(int)Names.PrimaryUse];
            developed.Text = namesList[(int)Names.Developed];
            checkedText.Text = namesList[(int)Names.Cheked];
            additionalField.Text = namesList[(int)Names.AdditionalField];
            additionalSurname.Text = namesList[(int)Names.AdditionalSurname];
            normalControl.Text = namesList[(int)Names.NormalControl];
            approved.Text = namesList[(int)Names.Approved];
            dateDeveloped.Text = namesList[(int)Names.DateDeveloped];
            dateCheked.Text = namesList[(int)Names.DateCheked];
            dateAdditionalSurname.Text = namesList[(int)Names.DateAdditionalSurname];
            dateNormalControl.Text = namesList[(int)Names.DateNormalControl];
            dateApproved.Text = namesList[(int)Names.DateApproved];
            letter1.Text = namesList[(int)Names.Letter1];
            letter2.Text = namesList[(int)Names.Letter2];
            letter3.Text = namesList[(int)Names.Letter3];
            refNumber.Text = namesList[(int)Names.RefNumber];
            originalInvNumber.Text = namesList[(int)Names.OriginalInvNumber];
            signDateOriginal.Text = namesList[(int)Names.SignDateOriginal];
            insteadInvNumber.Text = namesList[(int)Names.InsteadInvNumber];
            dublicateInvNumber.Text = namesList[(int)Names.DublicateInvNumber];
            signDateDublicate.Text = namesList[(int)Names.SignDateDublicate];
            //designationLU.Text = blank.DesignationLU;
            //approvalSheet.Text = blank.ApprovalSheet;
            //approvalDoc.Text = blank.ApprovalDoc;
            //customerIndex.Text = blank.CustomerIndex;
            firm.Text = namesList[(int)Names.Firm];
        }

        public void FillBlank()
        {
            Dictionary<int, string> namesList = new Dictionary<int, string>();
            namesList[(int)Names.Sheet] = designation.Text;
            namesList[(int)Names.Designation] = designation.Text;
            namesList[(int)Names.Name] = name.Text;
            namesList[(int)Names.Product] = product.Text;
            namesList[(int)Names.PrimaryUse] = primaryUse.Text;
            namesList[(int)Names.Developed] = developed.Text;
            namesList[(int)Names.Cheked] = checkedText.Text;
            namesList[(int)Names.AdditionalField] = additionalField.Text;
            namesList[(int)Names.AdditionalSurname] = additionalSurname.Text;
            namesList[(int)Names.NormalControl] = normalControl.Text;
            namesList[(int)Names.Approved] = approved.Text;
            namesList[(int)Names.DateDeveloped] = dateDeveloped.Text;
            namesList[(int)Names.DateCheked] = dateCheked.Text;
            namesList[(int)Names.DateAdditionalSurname] = dateAdditionalSurname.Text;
            namesList[(int)Names.DateNormalControl] = dateNormalControl.Text;
            namesList[(int)Names.DateApproved] = dateApproved.Text;
            namesList[(int)Names.Letter1] = letter1.Text;
            namesList[(int)Names.Letter2] = letter2.Text;
            namesList[(int)Names.Letter3] = letter3.Text;
            namesList[(int)Names.RefNumber] = refNumber.Text;
            namesList[(int)Names.OriginalInvNumber] = originalInvNumber.Text;
            namesList[(int)Names.SignDateOriginal] = signDateOriginal.Text;
            namesList[(int)Names.InsteadInvNumber] = insteadInvNumber.Text;
            namesList[(int)Names.DublicateInvNumber] = dublicateInvNumber.Text;
            namesList[(int)Names.SignDateDublicate] = signDateDublicate.Text;
            namesList[(int)Names.Firm] = firm.Text;

            blank.NamesList = namesList;
            blank.FillListPage();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            FillBlank();
            blank.FillListPage();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
