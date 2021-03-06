﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.View.Formatters
{
    class SpecFormatter : Formatter
    {
        protected override void GenerateFirstPage()
        {
            SpecFirstPage firstPage = new SpecFirstPage(sheet, pageCount);
            firstPage.format();
            currentRowData += firstPage.RowsCount;
            currentRowPage += firstPage.Height;
            pageNumber++;
        }

        protected override void GenerateSecondPage()
        {
            SpecSecondPage secondPage = new SpecSecondPage(sheet, currentRowPage, pageNumber);
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
