using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Blank
{
    public partial class BlankForm : Form
    {
        
        public BlankForm()
        {
            InitializeComponent();

        }

        private void BlankForm_Load(object sender, EventArgs e)
        {
            this.blankUC.FillTextBoxes();
        }


        private void SaveButton_Click(object sender, EventArgs e)
        {
            this.blankUC.FillBlank();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
