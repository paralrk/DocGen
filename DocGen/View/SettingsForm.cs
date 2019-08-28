using DocGen.Model;
using DocGen.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocGen.View
{
    public partial class SettingsForm : Form
    {
        SettingsFactory factory;
        Settings settings;
        public SettingsForm()
        {
            InitializeComponent();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            settings.Description = this.designatorBox.Text;
            settings.Type = typeBox.Text;
            settings.ManufacturerPartNumber = manufacturerPartNumberBox.Text;
            settings.Description = descriptionBox.Text;
            settings.Manufacturer = manufacturerBox.Text;
            settings.Note = noteBox.Text;
            settings.Note1 = note1Box.Text;
            settings.Quantity = quantityBox.Text;

            settings.GroupLimitPE3 = (int)groupLimitPE3UpDown.Value;
            settings.GroupLimitSpec = (int)groupLimitSpecUpDown.Value;
            settings.StartPositionNumber = (int)startPositionUpDown.Value;
            settings.PositionInc = (int)positionIncUpDown.Value;

            settings.MinPageForRegList = (int)regListUpDown.Value;

            JsonHelper.SerializeJson(settings);
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            factory = new SettingsFactory();
            settings = factory.GetSettings();
            FillForm(settings);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void DefaultButton_Click(object sender, EventArgs e)
        {
            settings.DefaultSettings();
            FillForm(settings);
        }

        private void FillForm(Settings settings)
        {
            this.designatorBox.Text = settings.Designator;
            this.typeBox.Text = settings.Type;
            this.manufacturerPartNumberBox.Text = settings.ManufacturerPartNumber;
            this.descriptionBox.Text = settings.Description;
            this.manufacturerBox.Text = settings.Manufacturer;
            this.noteBox.Text = settings.Note;
            this.note1Box.Text = settings.Note1;
            this.quantityBox.Text = settings.Quantity;

            this.groupLimitPE3UpDown.Value = settings.GroupLimitPE3;
            this.groupLimitSpecUpDown.Value = settings.GroupLimitSpec;
            this.startPositionUpDown.Value = settings.StartPositionNumber;
            this.positionIncUpDown.Value = settings.PositionInc;

            this.regListUpDown.Value = settings.MinPageForRegList;
        }
    }
}
