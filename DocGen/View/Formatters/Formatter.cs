using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using DocGen.Utils;
using DocGen.Model;

namespace DocGen.View.Formatters
{
    abstract class Formatter
    {
        protected Excel.Worksheet sheet;
        protected int currentRowData = 1;
        protected int currentRowPage = 1;
        protected int pageNumber = 1;
        protected int pageCount = 1;
        protected bool isRegistrationList = true;
        protected int minPageForRegList = 3;
        protected Settings settings;

        public Formatter ()
        {
            SettingsFactory factory = new SettingsFactory();
            this.settings = factory.GetSettings();
            this.minPageForRegList = settings.MinPageForRegList;
            this.sheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
        }
        public void Format()
        {
            Excel.Range usedRange = sheet.UsedRange;
            int usedRows = usedRange.Rows.Count;
            pageCount = PageCounter(usedRows);

            ExcelHelper.DisableUpdating();

            Debug.WriteLine("Formatting document");
            Stopwatch sw = new Stopwatch();
            sw.Start();
            GenerateFirstPage();
            sw.Stop();
            Debug.WriteLine("GenerateFirstPage() Elapsed={0}", sw.Elapsed);

            while (currentRowData < usedRows)
            {
                sw.Start();
                GenerateSecondPage();
                sw.Stop();
                Debug.WriteLine("GenerateSecondPage() Elapsed={0}", sw.Elapsed);
            }
            if (isRegistrationList && pageNumber > minPageForRegList)
            {
                sw.Start();
                GenerateRegistrationPage();
                sw.Stop();
                Debug.WriteLine("GenerateRegistrationPage() Elapsed={0}", sw.Elapsed);
            }

            ExcelHelper.SetPrintSettings(sheet);
            ExcelHelper.DisableZeros();
            ExcelHelper.EnableUpdating();
        }

        protected virtual void GenerateFirstPage()
        {

        }

        protected virtual void GenerateSecondPage()
        {

        }

        protected virtual void GenerateRegistrationPage()
        {
            RegistrationList regList = new RegistrationList(sheet, currentRowPage, pageNumber);
            regList.format();
            currentRowData += regList.RowsCount;
            currentRowPage += regList.Height;
            pageNumber++;
        }

        protected virtual int PageCounter(int usedRange)
        {
            return 1;
        }
    }
}
