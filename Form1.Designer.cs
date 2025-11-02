namespace PicExcleApp;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  设计器创建的DataGridView控件
    /// </summary>
    private System.Windows.Forms.DataGridView dataGridView;
    private System.Windows.Forms.Panel centerPanel;
    private System.Windows.Forms.PictureBox pictureBox;
    private System.Windows.Forms.ListBox imageListBox;
    private System.Windows.Forms.Panel topPanel;
    private System.Windows.Forms.Button importButton;
    private System.Windows.Forms.Button importCommunityButton;
    private System.Windows.Forms.Button importKeywordButton;
    private System.Windows.Forms.Button processButton;
    private System.Windows.Forms.Button exportButton;
    private System.Windows.Forms.ProgressBar progressBar;
    private System.Windows.Forms.RichTextBox logTextBox;
    private System.Windows.Forms.TableLayoutPanel mainLayoutPanel;

    /// <summary>
    ///  Clean up any resources being used.
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
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        this.components = new System.ComponentModel.Container();
        System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
        this.dataGridView = new System.Windows.Forms.DataGridView();
        this.centerPanel = new System.Windows.Forms.Panel();
        this.pictureBox = new System.Windows.Forms.PictureBox();
        this.imageListBox = new System.Windows.Forms.ListBox();
        this.topPanel = new System.Windows.Forms.Panel();
        this.exportButton = new System.Windows.Forms.Button();
        this.processButton = new System.Windows.Forms.Button();
        this.importKeywordButton = new System.Windows.Forms.Button();
        this.importCommunityButton = new System.Windows.Forms.Button();
        this.importButton = new System.Windows.Forms.Button();
        this.progressBar = new System.Windows.Forms.ProgressBar();
        this.logTextBox = new System.Windows.Forms.RichTextBox();
        this.mainLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
        // 移除单独定义的列对象，因为这些已经在AddRange中直接创建了
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
        this.centerPanel.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
        this.topPanel.SuspendLayout();
        this.mainLayoutPanel.SuspendLayout();
        this.SuspendLayout();
        // 
        // dataGridView
        // 
        this.dataGridView.AllowUserToAddRows = false;
        this.dataGridView.AllowUserToDeleteRows = false;
        this.dataGridView.AllowUserToOrderColumns = true;
        this.dataGridView.AllowUserToResizeColumns = true;
        this.dataGridView.AllowUserToResizeRows = false;
        this.dataGridView.AutoGenerateColumns = false;
        this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
        this.dataGridView.ColumnHeadersHeight = 80;
        this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "序号", HeaderText = "序号", MinimumWidth = 60, ReadOnly = true },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "CreateTime", HeaderText = "信访日期", DataPropertyName = "CreateTime", MinimumWidth = 100, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "WorkOrderNumber", HeaderText = "工单号", DataPropertyName = "WorkOrderNumber", MinimumWidth = 120, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "来源", HeaderText = "信访来源", DataPropertyName = "Source", MinimumWidth = 80, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "Category", HeaderText = "分类", DataPropertyName = "Category", MinimumWidth = 80, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "涉及企业", HeaderText = "涉及企业", DataPropertyName = "HeatingArea", MinimumWidth = 100, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "Content", HeaderText = "投诉内容", DataPropertyName = "Content", MinimumWidth = 200, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "Phone", HeaderText = "联系电话", DataPropertyName = "Phone", MinimumWidth = 100, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "测温温度", HeaderText = "测温温度", DataPropertyName = "Temperature", MinimumWidth = 80, ReadOnly = false },
        new System.Windows.Forms.DataGridViewTextBoxColumn { Name = "Status", HeaderText = "处理结果", DataPropertyName = "Result", MinimumWidth = 100, ReadOnly = false }
        });
        this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
        this.dataGridView.Location = new System.Drawing.Point(0, 0);
        this.dataGridView.Margin = new System.Windows.Forms.Padding(0);
        this.dataGridView.MultiSelect = false;
        this.dataGridView.Name = "dataGridView";
        this.dataGridView.Padding = new System.Windows.Forms.Padding(0, 10, 0, 0);
        this.dataGridView.ReadOnly = false;
        this.dataGridView.RowHeadersVisible = false;
        this.dataGridView.ScrollBars = System.Windows.Forms.ScrollBars.Both;
        this.dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
        this.dataGridView.Size = new System.Drawing.Size(500, 400);
        this.dataGridView.TabIndex = 0;
        this.dataGridView.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridView_CellBeginEdit);
        this.dataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellClick);
        this.dataGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellDoubleClick);
        this.dataGridView.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellEndEdit);
        this.dataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView_DataError);
        this.dataGridView.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView_KeyDown);
        this.dataGridView.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_RowEnter);
        this.dataGridView.SelectionChanged += new System.EventHandler(this.dataGridView_SelectionChanged);
        this.dataGridView.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView_RowPostPaint);
        // 
        // centerPanel
        // 
        this.centerPanel.Controls.Add(this.dataGridView);
        this.centerPanel.Dock = System.Windows.Forms.DockStyle.Fill;
        this.centerPanel.Location = new System.Drawing.Point(200, 0);
        this.centerPanel.Margin = new System.Windows.Forms.Padding(0);
        this.centerPanel.Name = "centerPanel";
        this.centerPanel.Size = new System.Drawing.Size(400, 400);
        this.centerPanel.TabIndex = 1;
        // 
        // pictureBox
        // 
        this.pictureBox.Dock = System.Windows.Forms.DockStyle.Fill;
        this.pictureBox.Location = new System.Drawing.Point(600, 0);
        this.pictureBox.Margin = new System.Windows.Forms.Padding(0);
        this.pictureBox.Name = "pictureBox";
        this.pictureBox.Size = new System.Drawing.Size(250, 400);
        this.pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
        this.pictureBox.TabIndex = 2;
        this.pictureBox.TabStop = false;
        // 
        // imageListBox
        // 
        this.imageListBox.Dock = System.Windows.Forms.DockStyle.Fill;
        this.imageListBox.Location = new System.Drawing.Point(0, 0);
        this.imageListBox.Margin = new System.Windows.Forms.Padding(0);
        this.imageListBox.Name = "imageListBox";
        this.imageListBox.Size = new System.Drawing.Size(200, 400);
        this.imageListBox.TabIndex = 0;
        // 
        // topPanel
        // 
        this.topPanel.BackColor = System.Drawing.Color.LightGray;
        this.topPanel.Controls.Add(this.exportButton);
        this.topPanel.Controls.Add(this.processButton);
        this.topPanel.Controls.Add(this.importKeywordButton);
        this.topPanel.Controls.Add(this.importCommunityButton);
        this.topPanel.Controls.Add(this.importButton);
        this.topPanel.Dock = System.Windows.Forms.DockStyle.Top;
        this.topPanel.Location = new System.Drawing.Point(0, 0);
        this.topPanel.Name = "topPanel";
        this.topPanel.Size = new System.Drawing.Size(850, 100);
        this.topPanel.TabIndex = 0;
        // 
        // exportButton
        // 
        this.exportButton.BackColor = System.Drawing.Color.LightBlue;
        this.exportButton.Location = new System.Drawing.Point(540, 20);
        this.exportButton.Name = "exportButton";
        this.exportButton.Size = new System.Drawing.Size(100, 40);
        this.exportButton.TabIndex = 4;
        this.exportButton.Text = "导出Excel";
        this.exportButton.UseVisualStyleBackColor = false;
        this.exportButton.Click += new System.EventHandler(this.ExportButton_Click);
        // 
        // processButton
        // 
        this.processButton.BackColor = System.Drawing.Color.LightGreen;
        this.processButton.Location = new System.Drawing.Point(420, 20);
        this.processButton.Name = "processButton";
        this.processButton.Size = new System.Drawing.Size(100, 40);
        this.processButton.TabIndex = 3;
        this.processButton.Text = "开始处理";
        this.processButton.UseVisualStyleBackColor = false;
        this.processButton.Click += new System.EventHandler(this.ProcessButton_Click);
        // 
        // importKeywordButton
        // 
        this.importKeywordButton.Location = new System.Drawing.Point(280, 20);
        this.importKeywordButton.Name = "importKeywordButton";
        this.importKeywordButton.Size = new System.Drawing.Size(120, 40);
        this.importKeywordButton.TabIndex = 2;
        this.importKeywordButton.Text = "导入关键词";
        this.importKeywordButton.UseVisualStyleBackColor = true;
        this.importKeywordButton.Click += new System.EventHandler(this.ImportKeywordButton_Click);
        // 
        // importCommunityButton
        // 
        this.importCommunityButton.Location = new System.Drawing.Point(140, 20);
        this.importCommunityButton.Name = "importCommunityButton";
        this.importCommunityButton.Size = new System.Drawing.Size(120, 40);
        this.importCommunityButton.TabIndex = 1;
        this.importCommunityButton.Text = "导入小区映射";
        this.importCommunityButton.UseVisualStyleBackColor = true;
        this.importCommunityButton.Click += new System.EventHandler(this.ImportCommunityButton_Click);
        // 
        // importButton
        // 
        this.importButton.Location = new System.Drawing.Point(20, 20);
        this.importButton.Name = "importButton";
        this.importButton.Size = new System.Drawing.Size(100, 40);
        this.importButton.TabIndex = 0;
        this.importButton.Text = "导入图片";
        this.importButton.UseVisualStyleBackColor = true;
        this.importButton.Click += new System.EventHandler(this.ImportButton_Click);
        // 
        // progressBar
        // 
        this.progressBar.Dock = System.Windows.Forms.DockStyle.Top;
        this.progressBar.Location = new System.Drawing.Point(0, 100);
        this.progressBar.Name = "progressBar1";
        this.progressBar.Size = new System.Drawing.Size(850, 30);
        this.progressBar.TabIndex = 1;
        // 
        // logTextBox
        // 
        this.logTextBox.Dock = System.Windows.Forms.DockStyle.Bottom;
        this.logTextBox.Location = new System.Drawing.Point(0, 530);
        this.logTextBox.Name = "logTextBox";
        this.logTextBox.ReadOnly = true;
        this.logTextBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both;
        this.logTextBox.Size = new System.Drawing.Size(850, 200);
        this.logTextBox.TabIndex = 2;
        this.logTextBox.Text = "";
        // 
        // mainLayoutPanel
        // 
        this.mainLayoutPanel.ColumnCount = 3;
        this.mainLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 200F));
        this.mainLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
        this.mainLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 250F));
        this.mainLayoutPanel.Controls.Add(this.imageListBox, 0, 0);
        this.mainLayoutPanel.Controls.Add(this.centerPanel, 1, 0);
        this.mainLayoutPanel.Controls.Add(this.pictureBox, 2, 0);
        this.mainLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
        this.mainLayoutPanel.Location = new System.Drawing.Point(0, 130);
        this.mainLayoutPanel.Name = "mainLayoutPanel";
        this.mainLayoutPanel.RowCount = 1;
        this.mainLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
        this.mainLayoutPanel.Size = new System.Drawing.Size(850, 400);
        this.mainLayoutPanel.TabIndex = 3;
        // 
        // 列已经在AddRange中定义，不需要重复定义
        // 
        // Form1
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(850, 730);
        this.Controls.Add(this.mainLayoutPanel);
        this.Controls.Add(this.logTextBox);
        this.Controls.Add(this.progressBar);
        this.Controls.Add(this.topPanel);
        this.Name = "Form1";
        this.Text = "图片批量录入输出Excel系统";
        this.AllowDrop = true;
        this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
        this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
        this.Resize += new System.EventHandler(this.Form1_Resize);
        this.Shown += new System.EventHandler(this.Form1_Shown);
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
        this.centerPanel.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
        this.topPanel.ResumeLayout(false);
        this.mainLayoutPanel.ResumeLayout(false);
        this.ResumeLayout(false);
    }

    #endregion
}
