namespace MSGTool_SolidWorks
{
    partial class MyForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyForm));
            this.MainMenu = new System.Windows.Forms.MenuStrip();
            this.AddFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.Export = new System.Windows.Forms.ToolStripMenuItem();
            this.SwDocType = new System.Windows.Forms.ToolStripComboBox();
            this.ExtendName = new System.Windows.Forms.ToolStripComboBox();
            this.LMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.Sub_Remove = new System.Windows.Forms.ToolStripMenuItem();
            this.Sub_Clear = new System.Windows.Forms.ToolStripMenuItem();
            this.Sub_SwOPen = new System.Windows.Forms.ToolStripMenuItem();
            this.Sub_AddAsm = new System.Windows.Forms.ToolStripMenuItem();
            this.Sub_Options = new System.Windows.Forms.ToolStripMenuItem();
            this.ListFile = new System.Windows.Forms.ListView();
            this.Bar = new System.Windows.Forms.ProgressBar();
            this.MainMenu.SuspendLayout();
            this.LMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainMenu
            // 
            this.MainMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.MainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddFolder,
            this.Export,
            this.SwDocType,
            this.ExtendName});
            this.MainMenu.Location = new System.Drawing.Point(0, 0);
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Padding = new System.Windows.Forms.Padding(7, 3, 0, 3);
            this.MainMenu.Size = new System.Drawing.Size(458, 31);
            this.MainMenu.TabIndex = 1;
            this.MainMenu.Text = "menuStrip1";
            // 
            // AddFolder
            // 
            this.AddFolder.Name = "AddFolder";
            this.AddFolder.Size = new System.Drawing.Size(80, 25);
            this.AddFolder.Text = "添加文件夹";
            this.AddFolder.Click += new System.EventHandler(this.AddFolder_Click);
            // 
            // Export
            // 
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(68, 25);
            this.Export.Text = "导出文件";
            this.Export.Click += new System.EventHandler(this.Export_Click);
            // 
            // SwDocType
            // 
            this.SwDocType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.SwDocType.DropDownWidth = 100;
            this.SwDocType.Items.AddRange(new object[] {
            "工程图",
            "零件",
            "钣金",
            "装配|零件"});
            this.SwDocType.Name = "SwDocType";
            this.SwDocType.Size = new System.Drawing.Size(140, 25);
            this.SwDocType.SelectedIndexChanged += new System.EventHandler(this.SwDocType_SelectedIndexChanged);
            // 
            // ExtendName
            // 
            this.ExtendName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ExtendName.Items.AddRange(new object[] {
            "Pdf",
            "Dxf",
            "Dwg",
            "Png"});
            this.ExtendName.Name = "ExtendName";
            this.ExtendName.Size = new System.Drawing.Size(121, 25);
            // 
            // Menu
            // 
            this.LMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.LMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.Sub_Remove,
            this.Sub_Clear,
            this.Sub_SwOPen,
            this.Sub_AddAsm,
            this.Sub_Options});
            this.LMenu.Name = "Menu";
            this.LMenu.Size = new System.Drawing.Size(176, 114);
            // 
            // Sub_Remove
            // 
            this.Sub_Remove.Name = "Sub_Remove";
            this.Sub_Remove.Size = new System.Drawing.Size(175, 22);
            this.Sub_Remove.Text = "移除选择项";
            this.Sub_Remove.Click += new System.EventHandler(this.Remove_Click);
            // 
            // Sub_Clear
            // 
            this.Sub_Clear.Name = "Sub_Clear";
            this.Sub_Clear.Size = new System.Drawing.Size(175, 22);
            this.Sub_Clear.Text = "清空列表";
            this.Sub_Clear.Click += new System.EventHandler(this.Clear_Click);
            // 
            // Sub_SwOPen
            // 
            this.Sub_SwOPen.Name = "Sub_SwOPen";
            this.Sub_SwOPen.Size = new System.Drawing.Size(175, 22);
            this.Sub_SwOPen.Text = "SolidWorks中打开";
            this.Sub_SwOPen.Click += new System.EventHandler(this.SwOPen_Click);
            // 
            // Sub_AddAsm
            // 
            this.Sub_AddAsm.Name = "Sub_AddAsm";
            this.Sub_AddAsm.Size = new System.Drawing.Size(175, 22);
            this.Sub_AddAsm.Text = "添加当前装配";
            this.Sub_AddAsm.Click += new System.EventHandler(this.AddAsm_Click);
            // 
            // Sub_Options
            // 
            this.Sub_Options.Name = "Sub_Options";
            this.Sub_Options.Size = new System.Drawing.Size(175, 22);
            this.Sub_Options.Text = "设置";
            this.Sub_Options.Click += new System.EventHandler(this.Options_Click);
            // 
            // ListFile
            // 
            this.ListFile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ListFile.HideSelection = false;
            this.ListFile.LabelWrap = false;
            this.ListFile.Location = new System.Drawing.Point(0, 31);
            this.ListFile.Name = "ListFile";
            this.ListFile.Size = new System.Drawing.Size(458, 370);
            this.ListFile.TabIndex = 2;
            this.ListFile.UseCompatibleStateImageBehavior = false;
            this.ListFile.View = System.Windows.Forms.View.Details;
            this.ListFile.DragDrop += new System.Windows.Forms.DragEventHandler(this.ListFile_DragDrop);
            this.ListFile.DragEnter += new System.Windows.Forms.DragEventHandler(this.ListFile_DragEnter);
            // 
            // Bar
            // 
            this.Bar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Bar.Location = new System.Drawing.Point(0, 378);
            this.Bar.Name = "Bar";
            this.Bar.Size = new System.Drawing.Size(458, 23);
            this.Bar.TabIndex = 3;
            // 
            // MyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(458, 401);
            this.ContextMenuStrip = this.LMenu;
            this.Controls.Add(this.Bar);
            this.Controls.Add(this.ListFile);
            this.Controls.Add(this.MainMenu);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.MainMenu;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MyForm";
            this.Text = "ExportTool";
            this.Load += new System.EventHandler(this.MyForm_Load);
            this.MainMenu.ResumeLayout(false);
            this.MainMenu.PerformLayout();
            this.LMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.MenuStrip MainMenu;
        private System.Windows.Forms.ToolStripMenuItem AddFolder;
        private System.Windows.Forms.ToolStripMenuItem Export;
        private System.Windows.Forms.ToolStripComboBox SwDocType;
        private System.Windows.Forms.ContextMenuStrip LMenu;
        private System.Windows.Forms.ToolStripMenuItem Sub_Remove;
        private System.Windows.Forms.ToolStripMenuItem Sub_Clear;
        private System.Windows.Forms.ToolStripComboBox ExtendName;
        private System.Windows.Forms.ToolStripMenuItem Sub_SwOPen;
        private System.Windows.Forms.ToolStripMenuItem Sub_AddAsm;
        private System.Windows.Forms.ListView ListFile;
        private System.Windows.Forms.ToolStripMenuItem Sub_Options;
        private System.Windows.Forms.ProgressBar Bar;
    }
}