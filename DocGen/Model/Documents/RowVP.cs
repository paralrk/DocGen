using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Documents
{
    class RowVP : IRow
    {
        public int RowNumber { get; set; }
        public string Name { get; set; }
        public string ProductCode { get; set; }
        public string Designation { get; set; }
        public string Supplier { get; set; }
        public string WhereItUsed { get; set; }
        public int QuantityProduct { get; set; }
        public int QuantitySet { get; set; }
        public int QuantityAdjustment { get; set; }
        public int QuantityTotal { get; set; }
        public string Note { get; set; }

        // constructors

        public RowVP()
        {
            this.RowNumber = 0;
            this.Name = "";
            this.ProductCode = "";
            this.Designation = "";
            this.Supplier = "";
            this.WhereItUsed = "";
            this.QuantityProduct = 0;
            this.QuantitySet = 0;
            this.QuantityAdjustment = 0;
            this.QuantityTotal = 0;
            this.Note = "";
        }

        public RowVP(int rowNumber,
                        string name,
                        string productCode,
                        string designation,
                        string supplier,
                        string whereItUsed,
                        int quantityProduct,
                        int quantitySet,
                        int quantityAdjustment,
                        int quantityTotal,
                        string note)
        {
            this.RowNumber = rowNumber;
            this.Name = name;
            this.ProductCode = productCode;
            this.Designation = designation;
            this.Supplier = supplier;
            this.WhereItUsed = whereItUsed;
            this.QuantityProduct = quantityProduct;
            this.QuantitySet = quantitySet;
            this.QuantityAdjustment = quantityAdjustment;
            this.QuantityTotal = quantityTotal;
            this.Note = note;
        }
    }
}
