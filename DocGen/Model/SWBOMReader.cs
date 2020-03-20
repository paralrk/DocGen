using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using DocGen.Model;
using DocGen.Utils;

namespace DocGen.Model
{
    class SWBOMReader
    {

        //private Component[] bom;
        private List<SWComponents> bom;
        bool isOrderSet = false;
        Dictionary<int, string> fieldsOrder = new Dictionary<int, string>();

        private int FORMAT;
        private int POSITION;
        private int DESIGNATION;
        private int NAME;
        private int QUANTITY;
        private int NOTE;
        private int DOCUMENT_SECTION;
        private int CLASS;
		private int GOST;
		private int SIZES_PARAMETRES;
		private int REPLACEMENT;

        private bool isFormat = false;
        private bool isPosition = false;
        private bool isDesignation = false;
        private bool isName = false;
        private bool isQuantity = false;
        private bool isNote = false;
		private bool isDocumentSection = false;
        private bool isClass = false;
        private bool isGost = false;
		private bool isSizesParametres = false;
		private bool isReplacement = false;

        Excel.Worksheet bomSheet;


        //public Component[] ReadBOM()
        public List<SWComponents> ReadBOM()
        {
            bomSheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            Excel.Range usedRange = bomSheet.UsedRange;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            int usedRows = usedRange.Rows.Count;
            if (usedRows > 1)
            {
                SetOrder();
                if (isOrderSet)
                {
                    bom = new List<SWComponents>(usedRows - 1);
                    for (int i = 2; i <= usedRows; i++)
                    {
                        AddComponent(i);
                    }
                } else
                {
                    bom = null;
                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            // isOrderSet == null, bom == null
            return bom;
        }

        public void SetOrder()
        {
            SettingsFactory factory = new SettingsFactory();
            Settings settings = factory.GetSettings();
            isOrderSet = false;

            Excel.Range titleRange = (Excel.Range)bomSheet.Rows[1];
            Excel.Range cells = (Excel.Range)titleRange.Cells;
            Excel.Range testCell = (Excel.Range)cells.Cells[1, 1];
            int usedColumns = bomSheet.UsedRange.Columns.Count;
            if (!(bool)testCell.MergeCells)
            {

                for (int i = 1; i <= usedColumns; i++)
                {
                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWFormat))
                    {
                        FORMAT = i;
                        isFormat  = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWPosition))
                    {
                        POSITION  = i;
                        isPosition  = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWDesignation))
                    {
                        DESIGNATION = i;
                        isDesignation  = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWName))
                    {
                        NAME = i;
                        isName  = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWQuantity))
                    {
                        QUANTITY = i;
                        isQuantity  = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWNote))
                    {
                        NOTE = i;
                        isNote = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWDocumentSection))
                    {
                        DOCUMENT_SECTION = i;
                        isDocumentSection = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWClass))
                    {
                        CLASS = i;
                        isClass  = true;
                        continue;
                    }
					
					if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWGost))
                    {
                        GOST = i;
                        isGost   = true;
                        continue;
                    }
					
					if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWSizesParametres))
                    {
                        SIZES_PARAMETRES = i;
                        isSizesParametres   = true;
                        continue;
                    }
					
					if ((cells[1, i] as Excel.Range).Value2.Equals(settings.SWReplacement))
                    {
                        REPLACEMENT = i;
                        isReplacement   = true;
                        continue;
                    }
                }

            }

            if (isDocumentSection && (isName || isSizesParametres))
            {
                isOrderSet = true;
            }
        }

        public void AddComponent(int row)
        {

            // make and add new Component in bom with order from setOrder()
            SWComponents c = new SWComponents();
            Excel.Range cells = (Excel.Range)bomSheet.Cells;

            if (isFormat)
            {
                c.Format = Convert.ToString((cells[row, FORMAT] as Excel.Range).Value2);
            }

            if (isPosition)
            {
                c.Position  = Convert.ToString((cells[row, POSITION] as Excel.Range).Value2);
            }

            if (isDesignation )
            {
                c.Designation  = Convert.ToString((cells[row, DESIGNATION] as Excel.Range).Value2);
            }

            if (isName )
            {
                c.Name = Convert.ToString((cells[row, NAME] as Excel.Range).Value2);
            }

            if (isNote)
            {
                c.Note = Convert.ToString((cells[row, NOTE] as Excel.Range).Value2);
            }

            if (isDocumentSection)
            {
                c.DocumentSection = Convert.ToString((cells[row, DOCUMENT_SECTION] as Excel.Range).Value2);
            }
			
			if (isClass)
            {
                c.Class = Convert.ToString((cells[row, CLASS] as Excel.Range).Value2);
            }
			
			if (isGost)
            {
                c.Gost = Convert.ToString((cells[row, GOST] as Excel.Range).Value2);
            }
			
			if (isSizesParametres)
            {
                c.SizesParametres = Convert.ToString((cells[row, SIZES_PARAMETRES] as Excel.Range).Value2);
            }
			
			if (isReplacement)
            {
                c.Replacement = Convert.ToString((cells[row, REPLACEMENT] as Excel.Range).Value2);
            }

            if (isQuantity)
            {
                string quantity = Convert.ToString((cells[row, QUANTITY] as Excel.Range).Value2);
                int q = 1;
                try
                {
                    q = Convert.ToInt32(quantity);
                }
                catch (FormatException)
                {
                    Console.WriteLine($"Unable to parse '{quantity}'");
                }
                catch (OverflowException)
                {
                    Console.WriteLine("{0} is outside the range of the Int32 type.", quantity);
                }
                if (q > 0)
                {
                    c.Quantity = q;
                }
                else
                {
                    c.Quantity = 1;
                }
            }

            bom.Add(c);
        } // addComponent
    }
}
