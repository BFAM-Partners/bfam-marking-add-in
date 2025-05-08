using Microsoft.Office.Tools.Ribbon;

namespace MarkingSheet
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.loadPositionsButton = this.Factory.CreateRibbonButton();
            this.sendMarksToSophisButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.loadCdsPositions = this.Factory.CreateRibbonButton();
            this.sendCdsMarksButton = this.Factory.CreateRibbonButton();
            this.CB_Marking = this.Factory.CreateRibbonGroup();
            this.loadCbPositionsButton = this.Factory.CreateRibbonButton();
            this.bloombergGroup = this.Factory.CreateRibbonGroup();
            this.bbgDesButton = this.Factory.CreateRibbonButton();
            this.bbgHvgButton = this.Factory.CreateRibbonButton();
            this.bbgG3Button = this.Factory.CreateRibbonButton();
            this.bbgG7Button = this.Factory.CreateRibbonButton();
            this.bbgG8Button = this.Factory.CreateRibbonButton();
            this.bbgG10Button = this.Factory.CreateRibbonButton();
            this.bbgDvdButton = this.Factory.CreateRibbonButton();
            this.bbgAswButton = this.Factory.CreateRibbonButton();
            this.bbgYasButton = this.Factory.CreateRibbonButton();
            this.bbgCvmButton = this.Factory.CreateRibbonButton();
            this.ovcvButton = this.Factory.CreateRibbonButton();
            this.allqButton = this.Factory.CreateRibbonButton();
            this.bvalButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.CB_Marking.SuspendLayout();
            this.bloombergGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.CB_Marking);
            this.tab1.Groups.Add(this.bloombergGroup);
            this.tab1.Label = "BFAM Credit";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.loadPositionsButton);
            this.group1.Items.Add(this.sendMarksToSophisButton);
            this.group1.Label = "Bond Marking";
            this.group1.Name = "group1";
            // 
            // loadPositionsButton
            // 
            this.loadPositionsButton.Image = ((System.Drawing.Image)(resources.GetObject("loadPositionsButton.Image")));
            this.loadPositionsButton.Label = "Load bond positions";
            this.loadPositionsButton.Name = "loadPositionsButton";
            this.loadPositionsButton.ShowImage = true;
            this.loadPositionsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadPositionsButton_Click);
            // 
            // sendMarksToSophisButton
            // 
            this.sendMarksToSophisButton.Image = ((System.Drawing.Image)(resources.GetObject("sendMarksToSophisButton.Image")));
            this.sendMarksToSophisButton.Label = "Send bond marks to Sophis";
            this.sendMarksToSophisButton.Name = "sendMarksToSophisButton";
            this.sendMarksToSophisButton.ShowImage = true;
            this.sendMarksToSophisButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sendMarksToSophisButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.loadCdsPositions);
            this.group2.Items.Add(this.sendCdsMarksButton);
            this.group2.Label = "CDS Marking";
            this.group2.Name = "group2";
            // 
            // loadCdsPositions
            // 
            this.loadCdsPositions.Image = ((System.Drawing.Image)(resources.GetObject("loadCdsPositions.Image")));
            this.loadCdsPositions.Label = "Load CDS positions";
            this.loadCdsPositions.Name = "loadCdsPositions";
            this.loadCdsPositions.ShowImage = true;
            this.loadCdsPositions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadCdsPositions_Click);
            // 
            // sendCdsMarksButton
            // 
            this.sendCdsMarksButton.Image = ((System.Drawing.Image)(resources.GetObject("sendCdsMarksButton.Image")));
            this.sendCdsMarksButton.Label = "Send CDS marks to Sophis";
            this.sendCdsMarksButton.Name = "sendCdsMarksButton";
            this.sendCdsMarksButton.ShowImage = true;
            this.sendCdsMarksButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sendCdsMarksButton_Click);
            // 
            // CB_Marking
            // 
            this.CB_Marking.Items.Add(this.loadCbPositionsButton);
            this.CB_Marking.Label = "CB Marking";
            this.CB_Marking.Name = "CB_Marking";
            // 
            // loadCbPositionsButton
            // 
            this.loadCbPositionsButton.Image = global::MarkingSheet.Properties.Resources.loadCdsPositions_Image1;
            this.loadCbPositionsButton.Label = "Load CB positions";
            this.loadCbPositionsButton.Name = "loadCbPositionsButton";
            this.loadCbPositionsButton.ShowImage = true;
            this.loadCbPositionsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadCbPositionsButton_Click);
            // 
            // bloombergGroup
            // 
            this.bloombergGroup.Items.Add(this.bbgDesButton);
            this.bloombergGroup.Items.Add(this.bbgHvgButton);
            this.bloombergGroup.Items.Add(this.bbgG3Button);
            this.bloombergGroup.Items.Add(this.bbgG7Button);
            this.bloombergGroup.Items.Add(this.bbgG8Button);
            this.bloombergGroup.Items.Add(this.bbgG10Button);
            this.bloombergGroup.Items.Add(this.bbgDvdButton);
            this.bloombergGroup.Items.Add(this.bbgAswButton);
            this.bloombergGroup.Items.Add(this.bbgYasButton);
            this.bloombergGroup.Items.Add(this.bbgCvmButton);
            this.bloombergGroup.Items.Add(this.ovcvButton);
            this.bloombergGroup.Items.Add(this.allqButton);
            this.bloombergGroup.Items.Add(this.bvalButton);
            this.bloombergGroup.Label = "Bloomberg";
            this.bloombergGroup.Name = "bloombergGroup";
            // 
            // bbgDesButton
            // 
            this.bbgDesButton.Image = ((System.Drawing.Image)(resources.GetObject("bbgDesButton.Image")));
            this.bbgDesButton.Label = "DES";
            this.bbgDesButton.Name = "bbgDesButton";
            this.bbgDesButton.ShowImage = true;
            this.bbgDesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgHvgButton
            // 
            this.bbgHvgButton.Image = ((System.Drawing.Image)(resources.GetObject("bbgHvgButton.Image")));
            this.bbgHvgButton.Label = "HVG";
            this.bbgHvgButton.Name = "bbgHvgButton";
            this.bbgHvgButton.ShowImage = true;
            this.bbgHvgButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgG3Button
            // 
            this.bbgG3Button.Image = ((System.Drawing.Image)(resources.GetObject("bbgG3Button.Image")));
            this.bbgG3Button.Label = "G3";
            this.bbgG3Button.Name = "bbgG3Button";
            this.bbgG3Button.ShowImage = true;
            this.bbgG3Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgG7Button
            // 
            this.bbgG7Button.Image = ((System.Drawing.Image)(resources.GetObject("bbgG7Button.Image")));
            this.bbgG7Button.Label = "G7";
            this.bbgG7Button.Name = "bbgG7Button";
            this.bbgG7Button.ShowImage = true;
            this.bbgG7Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgG8Button
            // 
            this.bbgG8Button.Image = ((System.Drawing.Image)(resources.GetObject("bbgG8Button.Image")));
            this.bbgG8Button.Label = "G8";
            this.bbgG8Button.Name = "bbgG8Button";
            this.bbgG8Button.ShowImage = true;
            this.bbgG8Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgG10Button
            // 
            this.bbgG10Button.Image = ((System.Drawing.Image)(resources.GetObject("bbgG10Button.Image")));
            this.bbgG10Button.Label = "G10";
            this.bbgG10Button.Name = "bbgG10Button";
            this.bbgG10Button.ShowImage = true;
            this.bbgG10Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgDvdButton
            // 
            this.bbgDvdButton.Image = ((System.Drawing.Image)(resources.GetObject("bbgDvdButton.Image")));
            this.bbgDvdButton.Label = "DVD";
            this.bbgDvdButton.Name = "bbgDvdButton";
            this.bbgDvdButton.ShowImage = true;
            this.bbgDvdButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgAswButton
            // 
            this.bbgAswButton.Image = ((System.Drawing.Image)(resources.GetObject("bbgAswButton.Image")));
            this.bbgAswButton.Label = "ASW";
            this.bbgAswButton.Name = "bbgAswButton";
            this.bbgAswButton.ShowImage = true;
            this.bbgAswButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgYasButton
            // 
            this.bbgYasButton.Image = ((System.Drawing.Image)(resources.GetObject("bbgYasButton.Image")));
            this.bbgYasButton.Label = "YAS";
            this.bbgYasButton.Name = "bbgYasButton";
            this.bbgYasButton.ShowImage = true;
            this.bbgYasButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bbgCvmButton
            // 
            this.bbgCvmButton.Image = ((System.Drawing.Image)(resources.GetObject("bbgCvmButton.Image")));
            this.bbgCvmButton.Label = "CVM";
            this.bbgCvmButton.Name = "bbgCvmButton";
            this.bbgCvmButton.ShowImage = true;
            this.bbgCvmButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // ovcvButton
            // 
            this.ovcvButton.Image = ((System.Drawing.Image)(resources.GetObject("ovcvButton.Image")));
            this.ovcvButton.Label = "OVCV";
            this.ovcvButton.Name = "ovcvButton";
            this.ovcvButton.ShowImage = true;
            this.ovcvButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // allqButton
            // 
            this.allqButton.Image = ((System.Drawing.Image)(resources.GetObject("allqButton.Image")));
            this.allqButton.Label = "ALLQ";
            this.allqButton.Name = "allqButton";
            this.allqButton.ShowImage = true;
            this.allqButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // bvalButton
            // 
            this.bvalButton.Image = ((System.Drawing.Image)(resources.GetObject("bvalButton.Image")));
            this.bvalButton.Label = "BVAL";
            this.bvalButton.Name = "bvalButton";
            this.bvalButton.ShowImage = true;
            this.bvalButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bbgButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.CB_Marking.ResumeLayout(false);
            this.CB_Marking.PerformLayout();
            this.bloombergGroup.ResumeLayout(false);
            this.bloombergGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadPositionsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sendMarksToSophisButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadCdsPositions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sendCdsMarksButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CB_Marking;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadCbPositionsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup bloombergGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgDesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgHvgButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgG3Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgG7Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgG8Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgG10Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgDvdButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgAswButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgYasButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bbgCvmButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ovcvButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton allqButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bvalButton;

        //internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshBBGValuesButton;
        //internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox autoBbgUpdatesRefresh;
        //internal Microsoft.Office.Tools.Ribbon.RibbonEditBox bbgRefreshRate;
        //internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
