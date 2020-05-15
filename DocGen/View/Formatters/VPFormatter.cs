using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.View.Formatters
{
    class VPFormatter : Formatter
    {
        protected override void GenerateFirstPage()
        {
            VPFirstPage firstPage = new VPFirstPage(sheet, pageCount);
            firstPage.Format();
            currentRowData += firstPage.RowsCount;
            currentRowPage += firstPage.Height;
            pageNumber++;
        }

        protected override void GenerateSecondPage()
        {
            VPSecondPage secondPage = new VPSecondPage(sheet, currentRowPage, pageNumber);
            secondPage.Format();
            currentRowData += secondPage.RowsCount;
            currentRowPage += secondPage.Height;
            pageNumber++;
        }

        protected override int PageCounter(int usedRange)
        {
            int pages = 1;
            if (usedRange > 23)
            {
                // 23 - count of rows at first list and +1 at the end
                // add 28 helps get rid of rounding 
                // (for int: 10 / 29 = 0,  (10 + 28) / 29 = 1)
                pages = (usedRange - 23 + 28) / 29 + 1;
            }
            if (isRegistrationList && pages > 3)
            {
                pages++;
            }
            return pages;
        }
    }
}
