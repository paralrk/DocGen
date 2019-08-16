using DocGen.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    [Serializable]
    class Settings
    {
        public string Designator { get; set; } = "Designator";
        public string Type { get; set; } = "ValueType";
        public string ManufacturerPartNumber { get; set; } = "ValueManufacturerPartNumber";
        public string Description { get; set; } = "ValueDescription";
        public string Manufacturer { get; set; } = "ValueManufacturer";
        public string Note { get; set; } = "Примечание";
        public string Note1 { get; set; } = "Примечание 1";
        public string Quantity { get; set; } = "Quantity";
        public int GroupLimitPE3 { get; set; } = 5;
        public int GroupLimitSpec { get; set; } = 3;
        public int StartPositionNumber { get; set; } = 5;
        public int PositionInc { get; set; } = 2;
        public int MaxNameLengthSpec { get; set; } = 28;
        public int MaxNoteLengthSpec { get; set; } = 14;
        public int MaxNameLengthPE3 { get; set; } = 50;
        public int LimitNoteRowSpec { get; set; } = 60;
        public int LimitNoteRowPE3 { get; set; } = 50;

        public int MinPageForRegList { get; set; } = 3;


        private static Settings instance;

        private Settings()
        {
            DefaultSettings();
        }
        public static Settings Instance()
        {
            if (instance == null)
                instance = new Settings();
            return instance;
        }

        public void DefaultSettings()
        {
            Designator = "Designator";
            Type = "ValueType";
            ManufacturerPartNumber = "ValueManufacturerPartNumber";
            Description = "ValueDescription";
            Manufacturer = "ValueManufacturer";
            Note = "Примечание";
            Note1 = "Примечание 1";
            Quantity = "Quantity";
            GroupLimitPE3 = 5;
            GroupLimitSpec = 3;
            StartPositionNumber = 5;
            PositionInc = 2;
            MaxNameLengthSpec = 28;
            MaxNoteLengthSpec = 14;
            MaxNameLengthPE3 = 50;
            LimitNoteRowSpec = 60;
            LimitNoteRowPE3 = 50;
            MinPageForRegList = 3;
        }
    }
}
