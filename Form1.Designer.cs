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
    private System.Windows.Forms.Button clearImagesButton;
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
        dataGridView = new DataGridView();
        centerPanel = new Panel();
        pictureBox = new PictureBox();
        imageListBox = new ListBox();
        topPanel = new Panel();
        exportButton = new Button();
        processButton = new Button();
        importKeywordButton = new Button();
        importCommunityButton = new Button();
        importButton = new Button();
        clearImagesButton = new Button();
        progressBar = new ProgressBar();
        logTextBox = new RichTextBox();
        mainLayoutPanel = new TableLayoutPanel();
        ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
        centerPanel.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)pictureBox).BeginInit();
        topPanel.SuspendLayout();
        mainLayoutPanel.SuspendLayout();
        SuspendLayout();
        // 
        // dataGridView
        // 
        dataGridView.AllowUserToAddRows = false;
        dataGridView.AllowUserToDeleteRows = false;
        dataGridView.AllowUserToOrderColumns = true;
        dataGridView.AllowUserToResizeRows = false;
        dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        dataGridView.ColumnHeadersHeight = 80;
        dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        dataGridView.Dock = DockStyle.Fill;
        dataGridView.EditMode = DataGridViewEditMode.EditOnEnter;
        dataGridView.Location = new Point(0, 0);
        dataGridView.Margin = new Padding(0);
        dataGridView.MultiSelect = false;
        dataGridView.Name = "dataGridView";
        dataGridView.RowHeadersVisible = false;
        dataGridView.RowHeadersWidth = 62;
        dataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
        dataGridView.Size = new Size(977, 400);
        dataGridView.TabIndex = 0;
        dataGridView.CellBeginEdit += dataGridView_CellBeginEdit;
        dataGridView.CellClick += dataGridView_CellClick;
        dataGridView.CellDoubleClick += dataGridView_CellDoubleClick;
        dataGridView.CellEndEdit += dataGridView_CellEndEdit;
        dataGridView.DataError += dataGridView_DataError;
        dataGridView.RowEnter += dataGridView_RowEnter;
        dataGridView.RowPostPaint += dataGridView_RowPostPaint;
        dataGridView.SelectionChanged += dataGridView_SelectionChanged;
        dataGridView.KeyDown += dataGridView_KeyDown;
        // 
        // centerPanel
        // 
        centerPanel.Controls.Add(dataGridView);
        centerPanel.Dock = DockStyle.Fill;
        centerPanel.Location = new Point(300, 0);
        centerPanel.Margin = new Padding(0);
        centerPanel.Name = "centerPanel";
        centerPanel.Size = new Size(977, 400);
        centerPanel.TabIndex = 1;
        // 
        // pictureBox
        // 
        pictureBox.Dock = DockStyle.Fill;
        pictureBox.Location = new Point(1277, 0);
        pictureBox.Margin = new Padding(0);
        pictureBox.Name = "pictureBox";
        pictureBox.Size = new Size(250, 400);
        pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
        pictureBox.TabIndex = 2;
        pictureBox.TabStop = false;
        // 
        // imageListBox
        // 
        imageListBox.Dock = DockStyle.Fill;
        imageListBox.HorizontalScrollbar = true;
        imageListBox.IntegralHeight = false;
        imageListBox.Location = new Point(0, 0);
        imageListBox.Margin = new Padding(0);
        imageListBox.Name = "imageListBox";
        imageListBox.Size = new Size(300, 400);
        imageListBox.TabIndex = 0;
        // 
        // topPanel
        // 
        topPanel.BackColor = Color.LightGray;
        topPanel.Controls.Add(exportButton);
        topPanel.Controls.Add(processButton);
        topPanel.Controls.Add(importKeywordButton);
        topPanel.Controls.Add(importCommunityButton);
        topPanel.Controls.Add(importButton);
        topPanel.Controls.Add(clearImagesButton);
        topPanel.Dock = DockStyle.Top;
        topPanel.Location = new Point(0, 0);
        topPanel.Name = "topPanel";
        topPanel.Size = new Size(1527, 100);
        topPanel.TabIndex = 0;
        // 
        // exportButton
        // 
        exportButton.BackColor = Color.LightBlue;
        exportButton.Location = new Point(614, 20);
        exportButton.Name = "exportButton";
        exportButton.Size = new Size(101, 40);
        exportButton.TabIndex = 4;
        exportButton.Text = "导出Excel";
        exportButton.UseVisualStyleBackColor = false;
        exportButton.Click += ExportButton_Click;
        // 
        // processButton
        // 
        processButton.BackColor = Color.LightGreen;
        processButton.Location = new Point(504, 20);
        processButton.Name = "processButton";
        processButton.Size = new Size(101, 40);
        processButton.TabIndex = 3;
        processButton.Text = "开始处理";
        processButton.UseVisualStyleBackColor = false;
        processButton.Click += ProcessButton_Click;
        // 
        // importKeywordButton
        // 
        importKeywordButton.Location = new Point(374, 20);
        importKeywordButton.Name = "importKeywordButton";
        importKeywordButton.Size = new Size(119, 40);
        importKeywordButton.TabIndex = 2;
        importKeywordButton.Text = "导入关键词";
        importKeywordButton.UseVisualStyleBackColor = true;
        importKeywordButton.Click += ImportKeywordButton_Click;
        // 
        // importCommunityButton
        // 
        importCommunityButton.Location = new Point(242, 20);
        importCommunityButton.Name = "importCommunityButton";
        importCommunityButton.Size = new Size(119, 40);
        importCommunityButton.TabIndex = 1;
        importCommunityButton.Text = "导入小区映射";
        importCommunityButton.UseVisualStyleBackColor = true;
        importCommunityButton.Click += ImportCommunityButton_Click;
        // 
        // importButton
        // 
        importButton.Location = new Point(20, 20);
        importButton.Name = "importButton";
        importButton.Size = new Size(101, 40);
        importButton.TabIndex = 0;
        importButton.Text = "导入图片";
        importButton.UseVisualStyleBackColor = true;
        importButton.Click += ImportButton_Click;
        // 
        // clearImagesButton
        // 
        clearImagesButton.Location = new Point(130, 20);
        clearImagesButton.Name = "clearImagesButton";
        clearImagesButton.Size = new Size(101, 40);
        clearImagesButton.TabIndex = 5;
        clearImagesButton.Text = "清除图片";
        clearImagesButton.UseVisualStyleBackColor = true;
        clearImagesButton.Click += ClearImagesButton_Click;
        // 
        // progressBar
        // 
        progressBar.Dock = DockStyle.Top;
        progressBar.Location = new Point(0, 100);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(1527, 30);
        progressBar.TabIndex = 1;
        // 
        // logTextBox
        // 
        logTextBox.Dock = DockStyle.Bottom;
        logTextBox.Location = new Point(0, 530);
        logTextBox.Name = "logTextBox";
        logTextBox.ReadOnly = true;
        logTextBox.Size = new Size(1527, 200);
        logTextBox.TabIndex = 2;
        logTextBox.Text = "";
        // 
        // mainLayoutPanel
        // 
        mainLayoutPanel.ColumnCount = 3;
        mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
        mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
        mainLayoutPanel.Controls.Add(imageListBox, 0, 0);
        mainLayoutPanel.Controls.Add(centerPanel, 1, 0);
        mainLayoutPanel.Controls.Add(pictureBox, 2, 0);
        mainLayoutPanel.Dock = DockStyle.Fill;
        mainLayoutPanel.Location = new Point(0, 130);
        mainLayoutPanel.Name = "mainLayoutPanel";
        mainLayoutPanel.RowCount = 1;
        mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        mainLayoutPanel.Size = new Size(1527, 400);
        mainLayoutPanel.TabIndex = 3;
        // 
        // Form1
        // 
        AllowDrop = true;
        AutoScaleDimensions = new SizeF(11F, 24F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1527, 730);
        Controls.Add(mainLayoutPanel);
        Controls.Add(logTextBox);
        Controls.Add(progressBar);
        Controls.Add(topPanel);
        Name = "Form1";
        Text = "图片批量录入输出Excel系统";
        Shown += Form1_Shown;
        DragDrop += Form1_DragDrop;
        DragEnter += Form1_DragEnter;
        Resize += Form1_Resize;
        ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
        centerPanel.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)pictureBox).EndInit();
        topPanel.ResumeLayout(false);
        mainLayoutPanel.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion
}
