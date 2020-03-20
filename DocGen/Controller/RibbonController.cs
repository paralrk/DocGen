using DocGen.View.Blank;
using DocGen.Model.Documents;
using DocGen.View.EmptyDocuments;
using DocGen.View.Formatters;
using DocGen.Model.Exporters;
using DocGen.View.Unformatters;
using DocGen.Utils;
using DocGen.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace DocGen.Controller
{
    public class RibbonController
    {
        private Excel.Range cell = null;
        private DocumentsFactory documentsfactory;
        private EmptyDocumentsFactory emptyDocumentsFactory;
        private FormattersFactory formattersFactory;
        private UnformattersFactory unformattersFactory;
        private SheetBordersDrawer drawer;
        private Excel.Worksheet sheet;

        IExporter pe3Exporter;
        IExporter specExporter;

        BlankForm blankForm;
        SettingsForm settingsForm;
        AboutWindow about;

        private bool isBordersEnabled = false;
        private enum States
        {
            Editing = 1,
            Formatted = 2,
            NotDocument = 3

        }

        public RibbonController()
        {
            this.documentsfactory = new DocumentsFactory();
            this.emptyDocumentsFactory = new EmptyDocumentsFactory();
            this.unformattersFactory = new UnformattersFactory();
            this.formattersFactory = new FormattersFactory();
            this.drawer = new SheetBordersDrawer();
            this.pe3Exporter = new PE3Exporter();
            this.specExporter = new SpecificationExporter();
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;
        }

        public void GeneratePE3()
        {
            IDocument document = documentsfactory.GetPE3Document();
            document.Generate();
            if (document.IsGenerated())
            {
                pe3Exporter.Export(document);
            }

        }

        public void GenerateSpecification()
        {
            IDocument document = documentsfactory.GetSpecificationDocument();
            document.Generate();
            if (document.IsGenerated())
            {
                specExporter.Export(document);
            }
        }

        public void GenerateSWSpecification()
        {
            IDocument document = documentsfactory.GetSWSpecificationDocument();
            document.Generate();
            if (document.IsGenerated())
            {
                specExporter.Export(document);
            }
        }

        public void NewPE3()
        {
            EmptyDocument emptyPE3 = emptyDocumentsFactory.GetPE3EmptyDocument();
            emptyPE3.NewDocument();
            emptyPE3.InitFormatCells();
            emptyPE3.Format();
        }

        public void NewSpecification()
        {
            EmptyDocument emptySpec = emptyDocumentsFactory.GetSpecificationEmptyDocument();
            emptySpec.NewDocument();
            emptySpec.Format();
            emptySpec.InitFormatCells();

        }


        public void EditDocument()
        {
            if (CheckState() == States.Formatted)
            {
                ExcelHelper.SetupNormalView();
                Unformatter unformatter = unformattersFactory.GetUnformatter();
                if (unformatter != null)
                {
                    unformatter.Unformat();
                }
            }

        }

        public void FormatDocument()
        {

            if (CheckState() == States.Editing)
            {
                Formatter formatter = formattersFactory.GetFormatter();
                if (formatter != null)
                {
                    Debug.WriteLine("FormatDocument method");
                    Stopwatch sw = new Stopwatch();

                    sw.Start();
                    drawer.DeleteSheetBorders();
                    sw.Stop();
                    Debug.WriteLine("DeleteSheetBorders() Elapsed={0}", sw.Elapsed);

                    sw.Reset();
                    sw.Start();
                    formatter.Format();
                    sw.Stop();
                    Debug.WriteLine("Format Elapsed={0}", sw.Elapsed);
                }
            }
            else
            {
                ExcelHelper.SetupNormalView();
            }
        }

        public void PreparePrintableDocument()
        {
            FormatDocument();
            ExcelHelper.SetupPageBreaksView();

        }

        public void OpenBlank()
        {
            blankForm = new BlankForm();
            blankForm.ShowDialog();
        }

        public void DrawSheetBorders()
        {
            if (CheckState() == States.Editing)
            {
                drawer.EnableSheetChangeEvent();
                drawer.DrawSheetBorders();
            }
            isBordersEnabled = true;
        }

        public void DisableSheetBorders()
        {
            if (CheckState() == States.Editing)
            {
                drawer.DisableSheetChangeEvent();
                drawer.DeleteSheetBorders();
            }
            isBordersEnabled = false;
        }

        private void Application_SheetActivate(object Sh)
        {
            //sheet = (Excel.Worksheet)Sh;
            if (SheetHelper.isExisting("Содержание"))
            {
                if (CheckState() == States.Editing)
                {
                    if (isBordersEnabled)
                    {
                        drawer.EnableSheetChangeEvent();
                        drawer.DrawSheetBorders();
                    }
                    else
                    {
                        drawer.DisableSheetChangeEvent();
                        drawer.DeleteSheetBorders();
                    }
                }
            }
        }

        private States CheckState()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string type = ListPage.GetDocumentType(sheet);
            if (String.IsNullOrEmpty(type))
            {
                return States.NotDocument;
            }

            cell = (Excel.Range)sheet.Cells[1, 1];
            if ((bool)cell.MergeCells)
            {
                return States.Formatted;
            }
            else
            {
                return States.Editing;
            }
        }
        public void OpenSettings()
        {
            settingsForm = new SettingsForm();
            settingsForm.Controller = this;
            settingsForm.ShowDialog();
        }

        public void ShowAbout()
        {
            about = new AboutWindow();
            about.ShowDialog();
        }

        public void SetColumnsWidth(int widthScale)
        {
            sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            CellsSizeManager cellsManager = new CellsSizeManager();
            cellsManager.SetColumnWidth(sheet, widthScale);
        }

        public void SetRowsHeight(int heightScale)
        {
            sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            CellsSizeManager cellsManager = new CellsSizeManager();
            cellsManager.SetRowsHeight(sheet, heightScale);
        }

    }
}
