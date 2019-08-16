using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DocGen.Controller;
using DocGen.Model;

namespace DocGen
{
    public partial class DocGenRibbon
    {
        RibbonController controller;
        Settings settings;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.controller = new RibbonController();
            SettingsFactory  factory = new SettingsFactory();
            this.settings = factory.GetSettings();
        }

        private void PE3Button_Click(object sender, RibbonControlEventArgs e)
        {

            controller.GeneratePE3();
        }

        private void SpecButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.GenerateSpecification();
        }


        private void NewPE3Button_Click(object sender, RibbonControlEventArgs e)
        {
            controller.NewPE3();
        }

        private void NewSpecificationButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.NewSpecification();
        }


        private void EditButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.EditDocument();
        }

        private void FormatButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.FormatDocument();
        }

        private void PrintableButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.PreparePrintableDocument();
        }


        private void BlankButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.OpenBlank();
        }

        private void SettingsButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.OpenSettings();
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            controller.ShowAbout();
        }

        private void SheetBordersCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            if (sheetBordersCheckBox.Checked)
            {
                controller.DrawSheetBorders();
            } else
            {
                controller.DisableSheetBorders();
            }
        }
    }
}
