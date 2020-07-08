using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.View.Formatters
{
    class D33_UDFormatter : Formatter
    {
        protected override void GenerateFirstPage()
        {
            D33_UDFirstPage firstPage = new D33_UDFirstPage(sheet, pageCount);
            firstPage.format();
            currentRowData += firstPage.RowsCount;
            currentRowPage += firstPage.Height;
            pageNumber++;
        }

    }
}
