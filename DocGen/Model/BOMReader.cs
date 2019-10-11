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
    class BOMReader
    {

        //private Component[] bom;
        private List<Components> bom;
        bool isOrderSet = false;
        Dictionary<int, string> fieldsOrder = new Dictionary<int, string>();

        private int DESIGNATOR;
        private int TYPE;
        private int MANUFACTURER_PARTNUMBER;
        private int DESCRIPTION;
        private int MANUFACTURER;
        private int NOTE;
        private int NOTE1;
        private int QUANTITY;

        private bool isDesignator = false;
        private bool isType = false;
        private bool isManufacturerPartNumber = false;
        private bool isDescription = false;
        private bool isManufacturer = false;
        private bool isNote = false;
        private bool isNote1 = false;
        private bool isQuantity = false;

        Excel.Worksheet bomSheet;


        //public Component[] ReadBOM()
        public List<Components> ReadBOM()
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
                    bom = new List<Components>(usedRows - 1);
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
                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Designator))
                    {
                        DESIGNATOR = i;
                        isDesignator = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Type))
                    {
                        TYPE = i;
                        isType = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.ManufacturerPartNumber))
                    {
                        MANUFACTURER_PARTNUMBER = i;
                        isManufacturerPartNumber = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Description))
                    {
                        DESCRIPTION = i;
                        isDescription = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Manufacturer))
                    {
                        MANUFACTURER = i;
                        isManufacturer = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Note))
                    {
                        NOTE = i;
                        isNote = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Note1))
                    {
                        NOTE1 = i;
                        isNote1 = true;
                        continue;
                    }

                    if ((cells[1, i] as Excel.Range).Value2.Equals(settings.Quantity))
                    {
                        QUANTITY = i;
                        isQuantity = true;
                        continue;
                    }
                }

            }

            if (isDesignator && isManufacturerPartNumber)
            {
                isOrderSet = true;
            }
        }

        public void AddComponent(int row)
        {

            // make and add new Component in bom with order from setOrder()
            Components c = new Components();
            Part p = new Part();
            Excel.Range cells = (Excel.Range)bomSheet.Cells;
            string designator = "";

            if (isDesignator)
            {
                //c.AddDesignator(Convert.ToString((cells[row, DESIGNATOR] as Excel.Range).Value2));
                designator = Convert.ToString((cells[row, DESIGNATOR] as Excel.Range).Value2);
            }

            if (isType)
            {
                p.Type = Convert.ToString((cells[row, TYPE] as Excel.Range).Value2);
            }

            if (isManufacturerPartNumber)
            {
                p.ManufacturerPartNumber = Convert.ToString((cells[row, MANUFACTURER_PARTNUMBER] as Excel.Range).Value2);
            }

            if (isDescription)
            {
                p.Description = Convert.ToString((cells[row, DESCRIPTION] as Excel.Range).Value2);
            }

            if (isManufacturer)
            {
                p.Manufacturer = Convert.ToString((cells[row, MANUFACTURER] as Excel.Range).Value2);
            }

            if (isNote)
            {
                p.Note = Convert.ToString((cells[row, NOTE] as Excel.Range).Value2);
            }

            if (isNote1)
            {
                p.Note1 = Convert.ToString((cells[row, NOTE1] as Excel.Range).Value2);
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
            string[] splitted = DesignatorsSplitter.SplitDesignators(designator);
            foreach (string s in splitted)
            {
                c = new Components(p);
                c.AddDesignator(s);
                c.Quantity = 1;
                bom.Add(c);
            }

        } // addComponent
    }
}
