namespace word插件
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SelectExcelButton = this.Factory.CreateRibbonButton();
            this.SelecWordlButton = this.Factory.CreateRibbonButton();
            this.GenerateCatalog = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.SelectExcelButton);
            this.group1.Items.Add(this.SelecWordlButton);
            this.group1.Items.Add(this.GenerateCatalog);
            this.group1.Label = "文件选择";
            this.group1.Name = "group1";
            // 
            // SelectExcelButton
            // 
            this.SelectExcelButton.Label = "选择Excel文件";
            this.SelectExcelButton.Name = "SelectExcelButton";
            this.SelectExcelButton.SuperTip = "选择制作的Excel模板文件";
            this.SelectExcelButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectExcelButton_Click);
            // 
            // SelecWordlButton
            // 
            this.SelecWordlButton.Label = "选择Word文件";
            this.SelecWordlButton.Name = "SelecWordlButton";
            this.SelecWordlButton.SuperTip = "选择制作的word模板文件";
            this.SelecWordlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectWordButton_Click);
            // 
            // GenerateCatalog
            // 
            this.GenerateCatalog.Label = "选择保存文件夹";
            this.GenerateCatalog.Name = "GenerateCatalog";
            this.GenerateCatalog.SuperTip = "处理完成的文件保存位置，默认在模板目录下";
            this.GenerateCatalog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateCatalog_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.dropDown1);
            this.group2.Label = "数据处理";
            this.group2.Name = "group2";
            // 
            // dropDown1
            // 
            this.dropDown1.Label = "excel拆分依据";
            this.dropDown1.Name = "dropDown1";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SelectExcelButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SelecWordlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateCatalog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
