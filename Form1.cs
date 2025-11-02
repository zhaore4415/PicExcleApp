using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using PicExcleApp.Models;
using PicExcleApp.Services;

namespace PicExcleApp
{
    public partial class Form1 : Form
    {
        // 服务实例
        private readonly ImageProcessingService _imageProcessingService;
        private readonly OCRService _ocrService;
        private ContentParserService _contentParserService;
        private readonly ExcelExportService _excelExportService;
        private readonly ConfigManagerService _configManagerService;
        
        // 配置和数据
        private KeywordConfig _keywordConfig;
        private List<ComplaintData> _processingDataList;
        
        // UI控件
        private DataGridView dataGridView;
        private Panel centerPanel;
        private PictureBox pictureBox;
        private ListBox imageListBox;
        
        public Form1()
        {
            InitializeComponent();
            
            // 初始化服务
            _imageProcessingService = new ImageProcessingService();
            _ocrService = new OCRService();
            _configManagerService = new ConfigManagerService();
            _keywordConfig = _configManagerService.LoadConfig();
            _contentParserService = new ContentParserService(_ocrService, _keywordConfig);
            _excelExportService = new ExcelExportService();
            
            _processingDataList = new List<ComplaintData>();
            
            // 关键修复：使用BindingList作为数据源，避免CurrencyManager索引问题
            var initialBindingList = new BindingList<ComplaintData>();
            _processingDataList = new List<ComplaintData>(initialBindingList);
            
            // 初始化UI
            InitializeUI();
        }
        
        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "图片批量录入输出Excel系统";
            this.Width = 1000;
            this.Height = 800;
            
            // 创建面板和控件
            CreateMainControls();
            
            // 支持拖放功能
            this.AllowDrop = true;
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;
        }
        
        private void CreateMainControls()
        {
            // 顶部操作面板
            Panel topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = Color.LightGray
            };
            
            // 导入图片按钮
            Button importButton = new Button
            {
                Text = "导入图片",
                Location = new Point(20, 20),
                Width = 100,
                Height = 40
            };
            importButton.Click += ImportButton_Click;
            
            // 导入小区映射按钮
            Button importCommunityButton = new Button
            {
                Text = "导入小区映射",
                Location = new Point(140, 20),
                Width = 120,
                Height = 40
            };
            importCommunityButton.Click += ImportCommunityButton_Click;
            
            // 导入关键词按钮
            Button importKeywordButton = new Button
            {
                Text = "导入关键词",
                Location = new Point(280, 20),
                Width = 120,
                Height = 40
            };
            importKeywordButton.Click += ImportKeywordButton_Click;
            
            // 处理按钮
            Button processButton = new Button
            {
                Text = "开始处理",
                Location = new Point(420, 20),
                Width = 100,
                Height = 40,
                BackColor = Color.LightGreen
            };
            processButton.Click += ProcessButton_Click;
            
            // 导出Excel按钮
            Button exportButton = new Button
            {
                Text = "导出Excel",
                Location = new Point(540, 20),
                Width = 100,
                Height = 40,
                BackColor = Color.LightBlue
            };
            exportButton.Click += ExportButton_Click;
            
            topPanel.Controls.AddRange(new Control[] { importButton, importCommunityButton, importKeywordButton, processButton, exportButton });
            
            // 进度条
            ProgressBar progressBar = new ProgressBar
            {
                Name = "progressBar1",
                Dock = DockStyle.Top,
                Height = 30,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            
            // 日志文本框
            RichTextBox logTextBox = new RichTextBox
            {
                Name = "logTextBox",
                Dock = DockStyle.Bottom,
                Height = 200,
                ReadOnly = true,
                Multiline = true,
                ScrollBars = RichTextBoxScrollBars.Both
            };
            
            // 创建主布局面板 - 使用TableLayoutPanel确保精确控制
            TableLayoutPanel mainLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(0),
                Margin = new Padding(0),
                BorderStyle = BorderStyle.None
            };
            
            // 设置列宽度 - 使用绝对宽度和百分比宽度组合
            mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200)); // 左侧图片列表
            mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));   // 中间数据表格
            mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250)); // 右侧图片预览
            
            // 图片列表 - 左侧固定
            imageListBox = new ListBox
            {
                Name = "imageListBox",
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.MultiExtended,
                BorderStyle = BorderStyle.Fixed3D
            };
            
            // 图片预览窗口 - 右侧固定
            pictureBox = new PictureBox
            {
                Name = "pictureBox",
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.Fixed3D
            };
            
            // 中间面板用于放置DataGridView
            centerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0),
                BorderStyle = BorderStyle.None
            };
            
            // 重新配置DataGridView，启用编辑功能同时保持稳定性
            dataGridView = new DataGridView
            {
                Name = "dataGridView",
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,  // 启用编辑功能
                
                // 关键修复：使用单元格选择模式，避免行选择带来的索引问题
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                
                // 基本显示设置
                RowHeadersVisible = false,
                ColumnHeadersVisible = true,
                ColumnHeadersHeight = 80,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                EnableHeadersVisualStyles = true,
                
                // 基本边框和滚动条
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = ScrollBars.Both,
                
                // 基本布局设置
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                
                // 启用列重排序，但禁用行操作以避免索引问题
                AllowUserToOrderColumns = true,
                AllowUserToResizeColumns = true,
                AllowUserToResizeRows = false,
                
                // 简化边距设置
                Padding = new Padding(0, 10, 0, 0)
            };
            
            // 设置列 - 使用相同的列定义
            dataGridView.Columns.AddRange(
                new DataGridViewTextBoxColumn { DataPropertyName = "WorkOrderNumber", HeaderText = "工单号", Name = "WorkOrderNumber", MinimumWidth = 80 },
                new DataGridViewTextBoxColumn { DataPropertyName = "Name", HeaderText = "姓名", Name = "Name", MinimumWidth = 60 },
                new DataGridViewTextBoxColumn { DataPropertyName = "Phone", HeaderText = "电话", Name = "Phone", MinimumWidth = 100 },
                new DataGridViewTextBoxColumn { DataPropertyName = "Content", HeaderText = "内容", Name = "Content", MinimumWidth = 150 },
                new DataGridViewTextBoxColumn { DataPropertyName = "Category", HeaderText = "分类", Name = "Category", MinimumWidth = 60 },
                new DataGridViewTextBoxColumn { DataPropertyName = "HeatingArea", HeaderText = "所属供热区域", Name = "HeatingArea", MinimumWidth = 100 },
                new DataGridViewTextBoxColumn { DataPropertyName = "Status", HeaderText = "状态", Name = "Status", MinimumWidth = 50 },
                new DataGridViewTextBoxColumn { DataPropertyName = "ErrorMessage", HeaderText = "错误信息", Name = "ErrorMessage", MinimumWidth = 120 },
                new DataGridViewTextBoxColumn { DataPropertyName = "CreateTime", HeaderText = "来件日期", Name = "CreateTime", MinimumWidth = 100 }
            );
            
            // 关键修复：添加全面的错误处理，抑制所有数据绑定错误
            dataGridView.DataError += (sender, e) =>
            {
                // 取消错误，防止传播
                e.Cancel = true;
                Log($"DataGridView数据错误已抑制: {e.Exception?.Message ?? "未知错误"}");
            };
            
            // 关键修复：处理所有可能导致索引越界的事件
            dataGridView.RowEnter += (sender, e) =>
            {
                // 阻止行进入事件，避免CurrencyManager索引问题
                // 清除选择状态，避免索引计算
                dataGridView.ClearSelection();
                dataGridView.CurrentCell = null;
            };
            
            // 添加单元格编辑事件处理
            dataGridView.CellEndEdit += (sender, e) =>
            {
                try
                {
                    // 获取编辑后的值并更新数据源
                    if (dataGridView.Rows[e.RowIndex].DataBoundItem is ComplaintData data)
                    {
                        var columnName = dataGridView.Columns[e.ColumnIndex].DataPropertyName;
                        var newValue = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                        
                        // 更新对应的数据属性
                        if (!string.IsNullOrEmpty(columnName) && newValue != null)
                        {
                            var property = typeof(ComplaintData).GetProperty(columnName);
                            if (property != null && property.CanWrite)
                            {
                                property.SetValue(data, Convert.ChangeType(newValue, property.PropertyType));
                                Log($"已更新第{e.RowIndex + 1}行的{columnName}字段");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"单元格编辑保存失败: {ex.Message}");
                }
            };
            
            dataGridView.SelectionChanged += (sender, e) =>
            {
                // 清除选择，避免索引问题
                if (dataGridView != null && !dataGridView.IsDisposed && dataGridView.Rows.Count > 0)
                {
                    dataGridView.ClearSelection();
                    dataGridView.CurrentCell = null;
                }
            };
            
            // 禁用鼠标事件，避免索引计算
            dataGridView.MouseDown += (sender, e) =>
            {
                // 只允许表头点击，阻止数据行点击
                var hitTest = dataGridView.HitTest(e.X, e.Y);
                if (hitTest.Type != DataGridViewHitTestType.ColumnHeader && hitTest.Type != DataGridViewHitTestType.None)
                {
                    // 清除选择以避免索引问题
                    dataGridView.ClearSelection();
                    dataGridView.CurrentCell = null;
                }
            };
            
            // 添加到中间面板
            centerPanel.Controls.Add(dataGridView);
            
            // 添加到主布局面板
            mainLayoutPanel.Controls.Add(imageListBox, 0, 0);
            mainLayoutPanel.Controls.Add(centerPanel, 1, 0);
            mainLayoutPanel.Controls.Add(pictureBox, 2, 0);
            
            // 添加到窗体
            this.Controls.AddRange(new Control[] { logTextBox, progressBar, mainLayoutPanel, topPanel });
            
            // 强制刷新布局的关键事件
            this.Shown += (sender, e) =>
            {
                this.BeginInvoke(new Action(() =>
                {
                    // 强制刷新整个控件层次
                    this.SuspendLayout();
                    
                    // 确保DataGridView正确显示
                    dataGridView.SuspendLayout();
                    dataGridView.ColumnHeadersHeight = 80;
                    dataGridView.ColumnHeadersVisible = true;
                    dataGridView.Padding = new Padding(0, 10, 0, 0);  // 再次确认顶部内边距
                    
                    // 强制布局更新
                    dataGridView.ResumeLayout(false);
                    dataGridView.PerformLayout();
                    
                    // 更新整个窗体
                    this.ResumeLayout(false);
                    this.PerformLayout();
                    
                    // 强制重绘DataGridView
                    dataGridView.Invalidate();
                    dataGridView.Refresh();
                }));
            };
            
            // 窗口大小改变事件
            this.Resize += (sender, e) =>
            {
                this.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        // 确保列头始终可见
                        if (dataGridView != null && !dataGridView.IsDisposed)
                        {
                            dataGridView.ColumnHeadersVisible = true;
                            dataGridView.ColumnHeadersHeight = 80;  // 保持高列头
                            dataGridView.Padding = new Padding(0, 10, 0, 0);  // 确保列头向下偏移
                            dataGridView.Refresh();
                            dataGridView.Update();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Resize事件异常: {ex.Message}");
                    }
                }));
            };
            
            // 设置数据绑定
            dataGridView.DataSource = _processingDataList;
        }
        
        private void Log(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;
                
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(Log), message);
                return;
            }
            
            RichTextBox logTextBox = this.Controls.Find("logTextBox", true).FirstOrDefault() as RichTextBox;
            if (logTextBox != null && !logTextBox.IsDisposed)
            {
                logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
                logTextBox.ScrollToCaret();
            }
        }
        
        private void UpdateProgress(int progress)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateProgress), progress);
                return;
            }
            
            ProgressBar progressBar = this.Controls.Find("progressBar1", true).FirstOrDefault() as ProgressBar;
            if (progressBar != null)
            {
                progressBar.Value = Math.Min(100, Math.Max(0, progress));
            }
        }
        
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }
        
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddImagesToProcess(files);
        }
        
        private void ImportButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp|所有文件|*.*";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    AddImagesToProcess(openFileDialog.FileNames);
                }
            }
        }
        
        private void AddImagesToProcess(string[]? files)
        {
            if (files == null || files.Length == 0 || _imageProcessingService == null)
                return;
                
            ListBox imageListBox = this.Controls.Find("imageListBox", true).FirstOrDefault() as ListBox;
            if (imageListBox == null || imageListBox.IsDisposed)
                return;
            
            int addedCount = 0;
            foreach (string file in files)
            {
                if (!string.IsNullOrEmpty(file) && _imageProcessingService.IsSupportedImageFormat(file))
                {
                    if (!imageListBox.Items.Contains(file))
                    {
                        imageListBox.Items.Add(file);
                        addedCount++;
                    }
                }
            }
            
            if (addedCount > 0)
            {
                Log($"已添加 {addedCount} 张图片");
            }
        }
        
        private async void ProcessButton_Click(object sender, EventArgs e)
        {
            ListBox imageListBox = this.Controls.Find("imageListBox", true).FirstOrDefault() as ListBox;
            if (imageListBox == null || imageListBox.Items.Count == 0)
            {
                MessageBox.Show("请先导入图片！");
                return;
            }
            
            _processingDataList.Clear();
            UpdateDataGridView();
            
            Log("开始处理图片...");
            await ProcessImagesAsync(imageListBox.Items.Cast<string>().ToArray());
            Log("图片处理完成");
        }
        
        private async Task ProcessImagesAsync(string[] imagePaths)
        {
            int totalImages = imagePaths.Length;
            
            await Task.Run(() =>
            {
                for (int i = 0; i < totalImages; i++)
                {
                    string imagePath = imagePaths[i];
                    try
                    {
                        Log($"处理图片 {i + 1}/{totalImages}: {Path.GetFileName(imagePath)}");
                        
                        // 预处理图片
                        var processedImage = _imageProcessingService.PreprocessImage(imagePath);
                        
                        // 文字识别
                        string ocrText = _ocrService.RecognizeText(processedImage);
                        
                        // 解析数据
                        var complaintData = _contentParserService.ParseToComplaintData(ocrText, imagePath);
                        
                        // 添加到结果列表
                        this.Invoke(new Action(() =>
                        {
                            _processingDataList.Add(complaintData);
                            UpdateDataGridView();
                        }));
                    }
                    catch (Exception ex)
                    {
                        Log($"处理失败: {ex.Message}");
                        
                        this.Invoke(new Action(() =>
                        {
                            _processingDataList.Add(new ComplaintData
                            {
                                OriginalImagePath = imagePath,
                                Status = "失败",
                                ErrorMessage = ex.Message
                            });
                            UpdateDataGridView();
                        }));
                    }
                    finally
                    {
                        UpdateProgress((i + 1) * 100 / totalImages);
                    }
                }
            });
        }
        
        private void UpdateDataGridView()
        {
            if (this.IsDisposed)
                return;
                
            DataGridView dataGridView = this.Controls.Find("dataGridView", true).FirstOrDefault() as DataGridView;
            if (dataGridView != null && !dataGridView.IsDisposed)
            {
                try
                {
                    // 关键修复：完全断开数据绑定，避免CurrencyManager索引问题
                    dataGridView.DataSource = null;
                    
                    // 暂停布局更新
                    dataGridView.SuspendLayout();
                    
                    // 清除所有状态
                    dataGridView.CurrentCell = null;
                    dataGridView.ClearSelection();
                    
                    // 使用BindingList避免直接绑定到List，提供更好的数据绑定支持
                    if (_processingDataList != null && _processingDataList.Count > 0)
                    {
                        var bindingList = new System.ComponentModel.BindingList<ComplaintData>(_processingDataList);
                        dataGridView.DataSource = bindingList;
                        
                        // 关键：不自动选择任何行，让用户手动选择
                        dataGridView.CurrentCell = null;
                        dataGridView.ClearSelection();
                    }
                    else
                    {
                        // 空列表时使用空的BindingList
                        dataGridView.DataSource = new System.ComponentModel.BindingList<ComplaintData>();
                    }
                    
                    // 恢复布局
                    dataGridView.ResumeLayout(false);
                    dataGridView.PerformLayout();
                    
                    // 再次确保没有选择任何单元格
                    this.BeginInvoke(new Action(() =>
                    {
                        if (dataGridView != null && !dataGridView.IsDisposed)
                        {
                            dataGridView.CurrentCell = null;
                            dataGridView.ClearSelection();
                        }
                    }));
                }
                catch (Exception ex)
                {
                    Log($"更新DataGridView时出错: {ex.Message}");
                    // 安全恢复
                    if (dataGridView != null && !dataGridView.IsDisposed)
                    {
                        try
                        {
                            dataGridView.DataSource = null;
                            dataGridView.CurrentCell = null;
                            dataGridView.ClearSelection();
                            dataGridView.ResumeLayout();
                        }
                        catch (Exception innerEx)
                        {
                            Log($"恢复DataGridView状态时出错: {innerEx.Message}");
                        }
                    }
                }
            }
        }
        
        private void ExportButton_Click(object sender, EventArgs e)
        {
            if (_processingDataList.Count == 0)
            {
                MessageBox.Show("没有数据可导出！");
                return;
            }
            
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel文件|*.xlsx";
                saveFileDialog.FileName = $"投诉数据_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _excelExportService.ExportToExcel(_processingDataList, saveFileDialog.FileName);
                        Log($"数据已导出到: {saveFileDialog.FileName}");
                        MessageBox.Show("导出成功！");
                    }
                    catch (Exception ex)
                    {
                        Log($"导出失败: {ex.Message}");
                        MessageBox.Show("导出失败: " + ex.Message);
                    }
                }
            }
        }
        
        private void ImportCommunityButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel文件|*.xlsx";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var communityMap = _excelExportService.ImportCommunityMap(openFileDialog.FileName);
                        _keywordConfig.CommunityToAreaMap = communityMap;
                        _configManagerService.SaveConfig(_keywordConfig);
                        
                        // 更新内容解析服务的配置
                        _contentParserService = new ContentParserService(_ocrService, _keywordConfig);
                        
                        Log($"成功导入 {communityMap.Count} 个小区映射");
                        MessageBox.Show("小区映射导入成功！");
                    }
                    catch (Exception ex)
                    {
                        Log($"导入失败: {ex.Message}");
                        MessageBox.Show("导入失败: " + ex.Message);
                    }
                }
            }
        }
        
        private void ImportKeywordButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel文件|*.xlsx";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var keywordMap = _excelExportService.ImportKeywordMap(openFileDialog.FileName);
                        _keywordConfig.KeywordToCategoryMap = keywordMap;
                        _configManagerService.SaveConfig(_keywordConfig);
                        
                        // 更新内容解析服务的配置
                        _contentParserService = new ContentParserService(_ocrService, _keywordConfig);
                        
                        Log($"成功导入 {keywordMap.Count} 个关键词配置");
                        MessageBox.Show("关键词配置导入成功！");
                    }
                    catch (Exception ex)
                    {
                        Log($"导入失败: {ex.Message}");
                        MessageBox.Show("导入失败: " + ex.Message);
                    }
                }
            }
        }
    }
}
