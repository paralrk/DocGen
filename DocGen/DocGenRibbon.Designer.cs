﻿namespace DocGen
{
    partial class DocGenRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DocGenRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.DocGenTab = this.Factory.CreateRibbonTab();
            this.DocumentsGroup = this.Factory.CreateRibbonGroup();
            this.PE3Button = this.Factory.CreateRibbonButton();
            this.SpecButton = this.Factory.CreateRibbonButton();
            this.VPButton = this.Factory.CreateRibbonButton();
            this.NewDocumentsGroup = this.Factory.CreateRibbonGroup();
            this.NewPE3Button = this.Factory.CreateRibbonButton();
            this.NewSpecificationButton = this.Factory.CreateRibbonButton();
            this.NewVPButton = this.Factory.CreateRibbonButton();
            this.NewD33_UD = this.Factory.CreateRibbonButton();
            this.EditGroup = this.Factory.CreateRibbonGroup();
            this.EditButton = this.Factory.CreateRibbonButton();
            this.FormatButton = this.Factory.CreateRibbonButton();
            this.PrintableButton = this.Factory.CreateRibbonButton();
            this.BlankButton = this.Factory.CreateRibbonButton();
            this.sheetBordersCheckBox = this.Factory.CreateRibbonCheckBox();
            this.SettingsGroup = this.Factory.CreateRibbonGroup();
            this.SettingsButton = this.Factory.CreateRibbonButton();
            this.AboutButton = this.Factory.CreateRibbonButton();
            this.SWDocumentsGroup = this.Factory.CreateRibbonGroup();
            this.SWSpecButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.DocGenTab.SuspendLayout();
            this.DocumentsGroup.SuspendLayout();
            this.NewDocumentsGroup.SuspendLayout();
            this.EditGroup.SuspendLayout();
            this.SettingsGroup.SuspendLayout();
            this.SWDocumentsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // DocGenTab
            // 
            this.DocGenTab.Groups.Add(this.DocumentsGroup);
            this.DocGenTab.Groups.Add(this.NewDocumentsGroup);
            this.DocGenTab.Groups.Add(this.EditGroup);
            this.DocGenTab.Groups.Add(this.SettingsGroup);
            this.DocGenTab.Groups.Add(this.SWDocumentsGroup);
            this.DocGenTab.Label = "DocGen";
            this.DocGenTab.Name = "DocGenTab";
            // 
            // DocumentsGroup
            // 
            this.DocumentsGroup.Items.Add(this.PE3Button);
            this.DocumentsGroup.Items.Add(this.SpecButton);
            this.DocumentsGroup.Items.Add(this.VPButton);
            this.DocumentsGroup.Label = "Документы из Altium Designer";
            this.DocumentsGroup.Name = "DocumentsGroup";
            // 
            // PE3Button
            // 
            this.PE3Button.Label = "BOM -> ПЭ3";
            this.PE3Button.Name = "PE3Button";
            this.PE3Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PE3Button_Click);
            // 
            // SpecButton
            // 
            this.SpecButton.Label = "BOM -> Спецификация";
            this.SpecButton.Name = "SpecButton";
            this.SpecButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SpecButton_Click);
            // 
            // VPButton
            // 
            this.VPButton.Label = "BOM -> Ведомость покупных изделий";
            this.VPButton.Name = "VPButton";
            this.VPButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.VPButton_Click);
            // 
            // NewDocumentsGroup
            // 
            this.NewDocumentsGroup.Items.Add(this.NewPE3Button);
            this.NewDocumentsGroup.Items.Add(this.NewSpecificationButton);
            this.NewDocumentsGroup.Items.Add(this.NewVPButton);
            this.NewDocumentsGroup.Items.Add(this.NewD33_UD);
            this.NewDocumentsGroup.Label = "Новые документы";
            this.NewDocumentsGroup.Name = "NewDocumentsGroup";
            // 
            // NewPE3Button
            // 
            this.NewPE3Button.Label = "Перечень элементов ПЭ3";
            this.NewPE3Button.Name = "NewPE3Button";
            this.NewPE3Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewPE3Button_Click);
            // 
            // NewSpecificationButton
            // 
            this.NewSpecificationButton.Label = "Спецификация";
            this.NewSpecificationButton.Name = "NewSpecificationButton";
            this.NewSpecificationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewSpecificationButton_Click);
            // 
            // NewVPButton
            // 
            this.NewVPButton.Label = "Ведомость покупных изделий";
            this.NewVPButton.Name = "NewVPButton";
            this.NewVPButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewVPButton_Click);
            // 
            // NewD33_UD
            // 
            this.NewD33_UD.Label = "Д33-УД";
            this.NewD33_UD.Name = "NewD33_UD";
            this.NewD33_UD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewD33_UD_Click);
            // 
            // EditGroup
            // 
            this.EditGroup.Items.Add(this.EditButton);
            this.EditGroup.Items.Add(this.FormatButton);
            this.EditGroup.Items.Add(this.PrintableButton);
            this.EditGroup.Items.Add(this.BlankButton);
            this.EditGroup.Items.Add(this.sheetBordersCheckBox);
            this.EditGroup.Label = "Редактирование";
            this.EditGroup.Name = "EditGroup";
            // 
            // EditButton
            // 
            this.EditButton.Label = "Редактировать";
            this.EditButton.Name = "EditButton";
            this.EditButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditButton_Click);
            // 
            // FormatButton
            // 
            this.FormatButton.Label = "Оформить";
            this.FormatButton.Name = "FormatButton";
            this.FormatButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FormatButton_Click);
            // 
            // PrintableButton
            // 
            this.PrintableButton.Label = "Оформить к печати";
            this.PrintableButton.Name = "PrintableButton";
            this.PrintableButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PrintableButton_Click);
            // 
            // BlankButton
            // 
            this.BlankButton.Label = "Основная надпись";
            this.BlankButton.Name = "BlankButton";
            this.BlankButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BlankButton_Click);
            // 
            // sheetBordersCheckBox
            // 
            this.sheetBordersCheckBox.Label = "Границы листов";
            this.sheetBordersCheckBox.Name = "sheetBordersCheckBox";
            this.sheetBordersCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SheetBordersCheckBox_Click);
            // 
            // SettingsGroup
            // 
            this.SettingsGroup.Items.Add(this.SettingsButton);
            this.SettingsGroup.Items.Add(this.AboutButton);
            this.SettingsGroup.Label = "Настройки";
            this.SettingsGroup.Name = "SettingsGroup";
            // 
            // SettingsButton
            // 
            this.SettingsButton.Label = "Настройки";
            this.SettingsButton.Name = "SettingsButton";
            this.SettingsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SettingsButton_Click);
            // 
            // AboutButton
            // 
            this.AboutButton.Label = "О надстройке";
            this.AboutButton.Name = "AboutButton";
            this.AboutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);
            // 
            // SWDocumentsGroup
            // 
            this.SWDocumentsGroup.Items.Add(this.SWSpecButton);
            this.SWDocumentsGroup.Label = "Документы из SolidWorks";
            this.SWDocumentsGroup.Name = "SWDocumentsGroup";
            this.SWDocumentsGroup.Visible = false;
            // 
            // SWSpecButton
            // 
            this.SWSpecButton.Label = "SW -> Специфицация";
            this.SWSpecButton.Name = "SWSpecButton";
            this.SWSpecButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SWSpecButton_Click);
            // 
            // DocGenRibbon
            // 
            this.Name = "DocGenRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.DocGenTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.DocGenTab.ResumeLayout(false);
            this.DocGenTab.PerformLayout();
            this.DocumentsGroup.ResumeLayout(false);
            this.DocumentsGroup.PerformLayout();
            this.NewDocumentsGroup.ResumeLayout(false);
            this.NewDocumentsGroup.PerformLayout();
            this.EditGroup.ResumeLayout(false);
            this.EditGroup.PerformLayout();
            this.SettingsGroup.ResumeLayout(false);
            this.SettingsGroup.PerformLayout();
            this.SWDocumentsGroup.ResumeLayout(false);
            this.SWDocumentsGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab DocGenTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DocumentsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PE3Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SpecButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup EditGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EditButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PrintableButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SettingsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BlankButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FormatButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup NewDocumentsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewPE3Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewSpecificationButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox sheetBordersCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SWDocumentsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SWSpecButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewVPButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VPButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewD33_UD;
    }

    partial class ThisRibbonCollection
    {
        internal DocGenRibbon Ribbon1
        {
            get { return this.GetRibbon<DocGenRibbon>(); }
        }
    }
}
