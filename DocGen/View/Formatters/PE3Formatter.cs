using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class PE3Formatter : Formatter
    {

        protected override void GenerateFirstPage()
        {
            PE3FirstPage firstPage = new PE3FirstPage(sheet, pageCount);
            firstPage.format();
            currentRowData += firstPage.RowsCount;
            currentRowPage += firstPage.Height;
            pageNumber++;
        }

        protected override void GenerateSecondPage()
        {
            PE3SecondPage secondPage = new PE3SecondPage(sheet, currentRowPage, pageNumber);
            secondPage.format();
            currentRowData += secondPage.RowsCount;
            currentRowPage += secondPage.Height;
            pageNumber++;
        }

        protected override int PageCounter(int usedRange)
        {
            int pages = 1;
            if (usedRange > 23)
            {
                // 24 - count of rows at first list and header
                // and +1 at the end add 28 helps get rid of rounding 
                // (for int: 10 / 29 = 0,  (10 + 28) / 29 = 1)
                pages = (usedRange - 24 + 28) / 29 + 1;
            }
            if (isRegistrationList && pages > 3)
            {
                pages++;
            }
            return pages;
        }
    }
}
