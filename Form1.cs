using System.ComponentModel;
using System.Data;
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
        private readonly DatabaseService _databaseService;

        // 配置和数据
        private readonly KeywordConfig _keywordConfig;
        private readonly List<ComplaintData> _processingDataList;
        private readonly BindingList<ComplaintData> _bindingList; // 单一BindingList实例

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
            _databaseService = new DatabaseService();
            try
            {
                Log("数据库服务初始化成功");
            }
            catch (Exception ex)
            {
                Log($"数据库服务初始化失败: {ex.Message}");
            }

            // 初始化数据列表
            _processingDataList = [];
            _bindingList = new BindingList<ComplaintData>(_processingDataList); // 初始化单一BindingList实例

            // 设置数据绑定
            SetDataBinding();

            // 配置imageListBox支持拖拽排序和删除
            if (imageListBox != null)
            {
                imageListBox.AllowDrop = true;
                imageListBox.MouseDown += ImageListBox_MouseDown;
                imageListBox.DragEnter += ImageListBox_DragEnter;
                imageListBox.DragOver += ImageListBox_DragOver;
                imageListBox.DragDrop += ImageListBox_DragDrop;
                imageListBox.KeyDown += ImageListBox_KeyDown;
                imageListBox.MouseClick += ImageListBox_MouseClick;
            }

            // 添加全局快捷键支持
            this.KeyPreview = true; // 允许表单接收键盘事件
            this.KeyDown += Form1_KeyDown; // 绑定键盘事件
        }

        // 全局键盘事件处理
        private void Form1_KeyDown(object? sender, KeyEventArgs e)
        {
            // Ctrl+Shift+C 快捷键清除所有图片
            if (e.Control && e.Shift && e.KeyCode == Keys.C)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;

                // 确认对话框
                if (MessageBox.Show("确定要清除所有图片吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    ClearAllImages();
                }
            }
        }

        private void SetDataBinding()
        {
            // 检查必要组件是否存在
            if (_bindingList == null || dataGridView == null)
                return;

            // 清除所有现有列
            dataGridView.Columns.Clear();

            // 设置AutoGenerateColumns为false，防止自动生成不需要的列
            dataGridView.AutoGenerateColumns = false;

            // 使用BindingSource作为中间层，提高数据绑定的稳定性和可编辑性
            BindingSource bindingSource = new()
            {
                DataSource = _bindingList
            };

            // 设置数据源
            dataGridView.DataSource = bindingSource;

            // 设置列标题和格式为用户期望的样式
            // 首先设置AutoSizeColumnsMode为None，以确保手动设置的列宽生效
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            UpdateColumnHeadersToChinese();
        }

        /// <summary>
        /// 设置DataGridView的列标题和格式为用户期望的样式
        /// </summary>
        private void UpdateColumnHeadersToChinese()
        {
            if (dataGridView == null || dataGridView.Columns == null)
                return;

            // 确保AutoGenerateColumns为false，防止自动生成不需要的列
            dataGridView.AutoGenerateColumns = false;

            // 首先清除所有现有列
            dataGridView.Columns.Clear();

            // 确保只保留用户当前显示的列，不包含OriginalImagePath、Status和ErrorMessage列
            // 定义列名映射（中文列名 -> 英文属性名）
            Dictionary<string, string> headerToPropertyMap = new()
            {
                { "信访日期", "CreateTime" },
                { "工单号", "WorkOrderNumber" },
                { "信访来源", "Source" },
                { "分类", "Category" },
                { "涉及企业", "HeatingArea" },
                { "投诉内容", "Content" },
                { "联系电话", "Phone" },
                { "测温温度", "Temperature" },
                { "处理结果", "Result" }
            };

            // 重新创建并添加所需的列，按照用户期望的顺序（只保留当前显示的列）
            List<string> columnOrder = new() { "序号", "信访日期", "工单号", "信访来源", "分类", "涉及企业", "投诉内容", "联系电话", "测温温度", "处理结果" };

            // 创建序号列
            DataGridViewTextBoxColumn serialColumn = new()
            {
                Name = "序号",
                HeaderText = "序号",
                DataPropertyName = string.Empty,
                ReadOnly = true,
                Width = 150
            };
            dataGridView.Columns.Add(serialColumn);

            // 为每个需要的列创建并设置属性
            foreach (string headerText in columnOrder.Skip(1)) // 跳过序号列
            {
                if (headerToPropertyMap.TryGetValue(headerText, out string propertyName))
                {
                    DataGridViewTextBoxColumn column = new()
                    {
                        Name = headerText,
                        HeaderText = headerText,
                        DataPropertyName = propertyName,
                        ReadOnly = false
                    };

                            // 设置列宽度
                    switch (headerText)
                    {
                        case "序号":
                            column.Width = 30;
                            break;
                        case "信访日期":
                            column.Width = 180;
                            break;
                        case "工单号":
                            column.Width = 230;
                            break;
                        case "信访来源":
                        case "分类":
                        case "涉及企业":
                            column.Width = 150;
                            break;
                        case "投诉内容":
                            column.Width = 400;
                            column.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                            break;
                        case "联系电话":
                        case "测温温度":
                            column.Width = 150;
                            break;
                        case "处理结果":
                            column.Width = 150;
                            column.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                            break;
                        default:
                            column.Width = 100;
                            break;
                    }

                    dataGridView.Columns.Add(column);
                }
            }

            // 设置行高和列标题高度
            dataGridView.RowTemplate.Height = 60;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dataGridView.ColumnHeadersHeight = 80;
        }

        private void Log(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            if (InvokeRequired)
            {
                _ = Invoke(new Action<string>(Log), message);
                return;
            }

            if (logTextBox != null && !logTextBox.IsDisposed)
            {
                logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
                logTextBox.ScrollToCaret();
            }
        }

        private void UpdateProgress(int value)
        {
            if (InvokeRequired)
            {
                _ = Invoke(new Action<int>(UpdateProgress), value);
                return;
            }

            if (progressBar != null && !progressBar.IsDisposed)
            {
                progressBar.Value = Math.Min(100, Math.Max(0, value));
            }
        }

        private void Form1_DragEnter(object? sender, DragEventArgs e)
        {
            if (e?.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Form1_DragDrop(object? sender, DragEventArgs e)
        {
            if (e?.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddImagesToProcess(files);
            }
        }

        private void ImportButton_Click(object? sender, EventArgs e)
        {
            using OpenFileDialog openFileDialog = new();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp|所有文件|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                AddImagesToProcess(openFileDialog.FileNames);
            }
        }

        // 清除所有图片按钮事件处理
        private void ClearImagesButton_Click(object? sender, EventArgs e)
        {
            // 确认对话框
            if (MessageBox.Show("确定要清除所有图片吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClearAllImages();
            }
        }

        // 添加清除所有图片的方法
        private void ClearAllImages()
        {
            try
            {
                if (imageListBox != null && !imageListBox.IsDisposed)
                {
                    int count = imageListBox.Items.Count;
                    if (count > 0)
                    {
                        imageListBox.Items.Clear();
                        Log($"已清除所有 {count} 张图片");

                        // 同时清除处理数据
                        if (_processingDataList != null)
                        {
                            _processingDataList.Clear();
                            if (_bindingList != null)
                            {
                                _bindingList.ResetBindings();
                                UpdateDataGridView();
                            }
                        }
                    }
                    else
                    {
                        Log("没有图片需要清除");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"清除图片时出错: {ex.Message}");
                MessageBox.Show($"清除图片失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 为了方便测试，在Form1构造函数中添加一个清除所有图片的快捷键
        // 可以通过按 Ctrl+Shift+C 来快速清除所有图片

        private void AddImagesToProcess(string[]? files)
        {
            if (files == null || files.Length == 0 || _imageProcessingService == null)
            {
                return;
            }

            if (imageListBox == null || imageListBox.IsDisposed)
            {
                return;
            }

            int addedCount = 0;
            foreach (string file in files)
            {
                if (!string.IsNullOrEmpty(file) && _imageProcessingService.IsSupportedImageFormat(file))
                {
                    if (!imageListBox.Items.Contains(file))
                    {
                        _ = imageListBox.Items.Add(file);
                        addedCount++;
                    }
                }
            }

            if (addedCount > 0)
            {
                Log($"已添加 {addedCount} 张图片");
            }
        }

        // 拖拽排序相关事件和变量
        private int _dragIndex = -1;

        private void ImageListBox_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && imageListBox != null && imageListBox.SelectedIndex != -1)
            {
                _dragIndex = imageListBox.SelectedIndex;
                imageListBox.DoDragDrop(imageListBox.SelectedItem, DragDropEffects.Move);
            }
        }

        private void ImageListBox_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.StringFormat) ?? false)
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void ImageListBox_DragOver(object? sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ImageListBox_DragDrop(object? sender, DragEventArgs e)
        {
            if (imageListBox == null || imageListBox.IsDisposed)
                return;

            int targetIndex = imageListBox.IndexFromPoint(imageListBox.PointToClient(new Point(e.X, e.Y)));

            if (_dragIndex != -1 && targetIndex != -1 && _dragIndex != targetIndex)
            {
                try
                {
                    // 获取要移动的项
                    object draggedItem = imageListBox.Items[_dragIndex];

                    // 移除并插入到新位置
                    imageListBox.Items.RemoveAt(_dragIndex);

                    // 调整插入位置（如果目标位置在拖拽项之后）
                    if (targetIndex > _dragIndex)
                    {
                        targetIndex--;
                    }

                    imageListBox.Items.Insert(targetIndex, draggedItem);

                    // 选中移动后的项
                    imageListBox.SelectedIndex = targetIndex;

                    Log("图片顺序已调整");
                }
                catch (Exception ex)
                {
                    Log($"调整图片顺序时出错: {ex.Message}");
                }
            }

            _dragIndex = -1;
        }

        // 删除图片功能
        private void ImageListBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && imageListBox != null && imageListBox.SelectedIndex != -1)
            {
                DeleteSelectedImages();
            }
        }

        private void ImageListBox_MouseClick(object? sender, MouseEventArgs e)
        {
            // 右键点击显示删除菜单
            if (e.Button == MouseButtons.Right && imageListBox != null)
            {
                try
                {
                    // 无论是否有选中项，都显示右键菜单
                    ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

                    // 如果有选中项，显示删除选中图片选项
                    if (imageListBox.SelectedIndex != -1)
                    {
                        ToolStripMenuItem deleteItem = new ToolStripMenuItem("删除选中图片");
                        deleteItem.Click += (s, ev) => DeleteSelectedImages();
                        contextMenuStrip.Items.Add(deleteItem);
                    }

                    // 总是显示清除所有图片选项
                    ToolStripMenuItem clearAllItem = new ToolStripMenuItem("清除所有图片");
                    clearAllItem.Click += (s, ev) =>
                    {
                        // 确认对话框
                        if (MessageBox.Show("确定要清除所有图片吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            ClearAllImages();
                        }
                    };
                    contextMenuStrip.Items.Add(clearAllItem);

                    // 显示菜单
                    contextMenuStrip.Show(imageListBox, e.Location);
                }
                catch (Exception ex)
                {
                    Log($"显示右键菜单时出错: {ex.Message}");
                }
            }
        }

        private void DeleteSelectedImages()
        {
            if (imageListBox == null || imageListBox.IsDisposed || imageListBox.SelectedIndices.Count == 0)
                return;

            try
            {
                // 记录要删除的图片路径
                List<string> removedFiles = new List<string>();
                foreach (int index in imageListBox.SelectedIndices)
                {
                    if (index >= 0 && index < imageListBox.Items.Count)
                    {
                        removedFiles.Add(imageListBox.Items[index].ToString());
                    }
                }

                // 从后往前删除，避免索引变化问题
                List<int> selectedIndices = new(imageListBox.SelectedIndices.Cast<int>());
                selectedIndices.Sort((a, b) => b.CompareTo(a));

                int deleteCount = 0;
                foreach (int index in selectedIndices)
                {
                    if (index >= 0 && index < imageListBox.Items.Count)
                    {
                        imageListBox.Items.RemoveAt(index);
                        deleteCount++;
                    }
                }

                // 清除与删除图片相关的处理数据
                if (_processingDataList != null && _processingDataList.Count > 0)
                {
                    for (int i = _processingDataList.Count - 1; i >= 0; i--)
                    {
                        if (removedFiles.Contains(_processingDataList[i].OriginalImagePath))
                        {
                            _processingDataList.RemoveAt(i);
                        }
                    }

                    // 通知BindingList更新
                    if (_bindingList != null)
                    {
                        _bindingList.ResetBindings();
                        UpdateDataGridView();
                    }
                }

                // 显示成功消息
                if (deleteCount > 0)
                {
                    MessageBox.Show($"成功删除 {deleteCount} 张图片", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除图片失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void ProcessButton_Click(object? sender, EventArgs e)
        {
            if (imageListBox == null || imageListBox.Items.Count == 0)
            {
                _ = MessageBox.Show("请先导入图片！");
                return;
            }

            _processingDataList.Clear();
            UpdateDataGridView();

            Log("开始处理图片...");
            await ProcessImagesAsync([.. imageListBox.Items.Cast<string>()]);
            Log("图片处理完成");
        }

        private async Task ProcessImagesAsync(string[] imagePaths)
        {
            int totalImages = imagePaths?.Length ?? 0;

            await Task.Run(() =>
            {
                if (imagePaths == null || _processingDataList == null)
                {
                    return;
                }

                for (int i = 0; i < totalImages; i++)
                {
                    string imagePath = imagePaths[i];
                    try
                    {
                        Log($"处理图片 {i + 1}/{totalImages}: {Path.GetFileName(imagePath)}");

                        // 预处理图片
                        MemoryStream? processedImage = _imageProcessingService?.PreprocessImage(imagePath);

                        // 文字识别
                        string ocrText = _ocrService?.RecognizeText(processedImage);

                        // 解析数据
                        ComplaintData? complaintData = _contentParserService?.ParseToComplaintData(ocrText, imagePath);
                        if (complaintData != null)
                        {
                            // 添加到结果列表
                            Invoke(new Action(() =>
                              {
                                  _processingDataList.Add(complaintData);
                                  _bindingList.ResetBindings();
                                  UpdateDataGridView();
                              }));
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"处理失败: {ex.Message}");

                        Invoke(new Action(() =>
                          {
                              _processingDataList.Add(new ComplaintData
                              {
                                  OriginalImagePath = imagePath,
                                  Status = "失败",
                                  ErrorMessage = ex.Message
                              });
                              _bindingList.ResetBindings();
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
            if (IsDisposed || dataGridView == null || dataGridView.IsDisposed)
            {
                return;
            }

            try
            {
                // 使用UI线程安全的方式更新DataGridView
                if (InvokeRequired)
                {
                    Invoke(new Action(UpdateDataGridViewSafe));
                }
                else
                {
                    UpdateDataGridViewSafe();
                }
            }
            catch (Exception ex)
            {
                Log($"更新DataGridView时出错: {ex.Message}");
                // 安全恢复
                if (!IsDisposed && dataGridView != null && !dataGridView.IsDisposed)
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

        // 安全更新DataGridView的私有方法
        private void UpdateDataGridViewSafe()
        {
            try
            {
                // 全面的null检查
                if (dataGridView == null || dataGridView.IsDisposed)
                {
                    Log("DataGridView为null或已释放，无法更新");
                    return;
                }

                // 暂停布局更新以提高性能
                dataGridView.SuspendLayout();

                // 确保DataGridView是可编辑的
                dataGridView.ReadOnly = false;

                // 刷新BindingList以更新数据
                _bindingList?.ResetBindings();

                // 确保列标题保持中文
                // 首先设置AutoSizeColumnsMode为None，以确保手动设置的列宽生效
                dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                UpdateColumnHeadersToChinese();

                // 确保所有编辑相关设置正确
                dataGridView.Enabled = true;
                dataGridView.EditMode = DataGridViewEditMode.EditOnEnter;

                // 为序号列添加自动生成功能
                if (dataGridView.Rows != null && dataGridView.Columns != null && dataGridView.Columns.Contains("序号"))
                {
                    try
                    {
                        int serialColumnIndex = dataGridView.Columns["序号"].Index;
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            if (row != null && row.Cells != null && serialColumnIndex >= 0 && row.Cells.Count > serialColumnIndex)
                            {
                                DataGridViewCell cell = row.Cells[serialColumnIndex];
                                if (cell != null)
                                {
                                    cell.Value = row.Index + 1;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"设置序号列值时出错: {ex.Message}");
                    }
                }

                // 设置列的可编辑性
                if (dataGridView.Columns != null)
                {
                    foreach (DataGridViewColumn column in dataGridView.Columns)
                    {
                        if (column != null)
                        {
                            if (column.Name == "序号")
                            {
                                column.ReadOnly = true; // 序号列设为只读
                            }
                            else
                            {
                                column.ReadOnly = false; // 其他列保持可编辑
                            }
                        }
                    }
                }

                // 刷新以确保显示最新数据
                dataGridView.Refresh();

                // 恢复布局更新
                dataGridView.ResumeLayout();

                // 安全的日志记录
                try
                {
                    Log($"DataGridView已更新，当前行数: {dataGridView.Rows?.Count ?? 0}");
                }
                catch { }
            }
            catch (Exception ex)
            {
                Log($"安全更新DataGridView时出错: {ex.Message}");
            }
        }

        private void ExportButton_Click(object? sender, EventArgs e)
        {
            if (_processingDataList.Count == 0)
            {
                _ = MessageBox.Show("没有数据可导出！");
                return;
            }

            using SaveFileDialog saveFileDialog = new();
            saveFileDialog.Filter = "Excel文件|*.xlsx";
            saveFileDialog.FileName = $"投诉数据_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _excelExportService.ExportToExcel(_processingDataList, saveFileDialog.FileName);
                    Log($"数据已导出到: {saveFileDialog.FileName}");
                    _ = MessageBox.Show("导出成功！");
                }
                catch (Exception ex)
                {
                    Log($"导出失败: {ex.Message}");
                    _ = MessageBox.Show("导出失败: " + ex.Message);
                }
            }
        }

        private void ImportCommunityButton_Click(object? sender, EventArgs e)
        {
            using OpenFileDialog openFileDialog = new();
            openFileDialog.Filter = "Excel文件|*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Dictionary<string, string> communityMap = _excelExportService.ImportCommunityMap(openFileDialog.FileName);
                    _keywordConfig.CommunityToAreaMap = communityMap;
                    _configManagerService.SaveConfig(_keywordConfig);

                    // 更新内容解析服务的配置
                    _contentParserService = new ContentParserService(_ocrService, _keywordConfig);

                    Log($"成功导入 {communityMap.Count} 个小区映射");
                    _ = MessageBox.Show("小区映射导入成功！");
                }
                catch (Exception ex)
                {
                    Log($"导入失败: {ex.Message}");
                    _ = MessageBox.Show("导入失败: " + ex.Message);
                }
            }
        }

        private void ImportKeywordButton_Click(object? sender, EventArgs e)
        {
            using OpenFileDialog openFileDialog = new();
            openFileDialog.Filter = "Excel文件|*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Dictionary<string, string> keywordMap = _excelExportService.ImportKeywordMap(openFileDialog.FileName);
                    _keywordConfig.KeywordToCategoryMap = keywordMap;
                    _configManagerService.SaveConfig(_keywordConfig);

                    // 更新内容解析服务的配置
                    _contentParserService = new ContentParserService(_ocrService, _keywordConfig);

                    Log($"成功导入 {keywordMap.Count} 个关键词配置");
                    _ = MessageBox.Show("关键词配置导入成功！");
                }
                catch (Exception ex)
                {
                    Log($"导入失败: {ex.Message}");
                    _ = MessageBox.Show("导入失败: " + ex.Message);
                }
            }
        }

        // 事件处理方法
        private void dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (sender is DataGridView dgv && !dgv.ReadOnly)
                {
                    _ = dgv.BeginEdit(true);
                    Log($"双击单元格[{e.RowIndex},{e.ColumnIndex}]开始编辑");
                }
            }
        }

        private void dataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // 确保序号列在每次行绘制后都更新，添加全面的安全检查
            if (sender is DataGridView dgv && dgv.Columns != null && dgv.Columns.Contains("序号"))
            {
                try
                {
                    if (e.RowIndex >= 0 && e.RowIndex < dgv.Rows.Count && dgv.Rows[e.RowIndex] != null)
                    {
                        int serialColumnIndex = dgv.Columns["序号"].Index;
                        if (serialColumnIndex >= 0 && dgv.Rows[e.RowIndex].Cells.Count > serialColumnIndex)
                        {
                            DataGridViewCell cell = dgv.Rows[e.RowIndex].Cells[serialColumnIndex];
                            if (cell != null)
                            {
                                cell.Value = e.RowIndex + 1;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"更新序号列时出错: {ex.Message}");
                }
            }
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (sender is DataGridView dgv && dgv.CurrentCell != null && !dgv.ReadOnly)
            {
                // 按下任何字母或数字键直接开始编辑
                if (e.KeyCode is (>= Keys.A and <= Keys.Z) or
                    (>= Keys.D0 and <= Keys.D9) or
                    Keys.Delete or Keys.Back)
                {
                    _ = dgv.BeginEdit(true);
                    if (dgv.CurrentCell != null)
                    {
                        Log($"键盘操作开始编辑单元格[{dgv.CurrentCell.RowIndex},{dgv.CurrentCell.ColumnIndex}]");
                    }
                }
            }
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (sender is DataGridView dgv && !dgv.ReadOnly)
                {
                    // 点击单元格时立即开始编辑
                    _ = dgv.BeginEdit(true);
                    Log($"用户点击单元格[{e.RowIndex},{e.ColumnIndex}]开始编辑");
                }
            }
        }

        private void dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // 取消错误，防止传播
            e.Cancel = true;
            Log($"DataGridView数据错误已抑制: {e.Exception?.Message ?? "未知错误"}");
        }

        private void dataGridView_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            // 仅在有有效数据时允许选择，但不阻止编辑
            if (dataGridView.Rows.Count > 0 && e.RowIndex >= 0 && e.RowIndex < dataGridView.Rows.Count)
            {
                // 确保单元格在选择时可以被编辑
                dataGridView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            }
        }

        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // 确保编辑开始前单元格有效
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView.Rows.Count &&
                      e.ColumnIndex >= 0 && e.ColumnIndex < dataGridView.Columns.Count)
            {
                // 确保单元格在编辑前是可编辑的
                if (sender is DataGridView dgv)
                {
                    dgv.ReadOnly = false;
                    if (dgv.CurrentCell != null)
                    {
                        dgv.CurrentCell.ReadOnly = false;
                    }
                    if (e.ColumnIndex >= 0 && e.ColumnIndex < dgv.Columns.Count)
                    {
                        dgv.Columns[e.ColumnIndex].ReadOnly = false;
                    }
                    if (e.RowIndex >= 0 && e.RowIndex < dgv.Rows.Count)
                    {
                        dgv.Rows[e.RowIndex].ReadOnly = false;
                    }
                }

                // 记录开始编辑的单元格
                if (dataGridView != null && e.ColumnIndex >= 0 && e.ColumnIndex < dataGridView.Columns.Count)
                {
                    Log($"开始编辑单元格: 第{e.RowIndex + 1}行, {dataGridView.Columns[e.ColumnIndex].HeaderText}");
                }
            }
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // 获取编辑后的值并更新数据源
                if (dataGridView.Rows[e.RowIndex].DataBoundItem is ComplaintData data)
                {
                    string columnName = dataGridView.Columns[e.ColumnIndex].DataPropertyName;
                    string displayColumnName = dataGridView.Columns[e.ColumnIndex].Name; // 获取显示的列名
                    object? newValue = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                    // 检查是否编辑了投诉内容列，如果是，则重新匹配供热区域
                    bool isComplaintContentColumn = displayColumnName == "投诉内容" ||
                                                   dataGridView.Columns[e.ColumnIndex].HeaderText == "投诉内容";
                    string? originalHeatingArea = data.HeatingArea; // 保存原始供热区域值

                    // 获取原始值用于日志
                    object? oldValue = null;
                    System.Reflection.PropertyInfo? property = typeof(ComplaintData).GetProperty(columnName);

                    // 处理特殊列的映射（这些列没有DataPropertyName但需要手动映射）
                    bool isSpecialColumn = false;
                    switch (displayColumnName)
                    {
                        case "来源": // 信访来源
                            oldValue = data.Source;
                            data.Source = newValue?.ToString();
                            Log($"已成功更新第{e.RowIndex + 1}行的信访来源字段: {oldValue} -> {newValue}");
                            isSpecialColumn = true;
                            break;
                        case "测温温度":
                            oldValue = data.Temperature;
                            data.Temperature = newValue?.ToString();
                            Log($"已成功更新第{e.RowIndex + 1}行的测温温度字段: {oldValue} -> {newValue}");
                            isSpecialColumn = true;
                            break;
                        case "处理结果":
                            oldValue = data.Result;
                            data.Result = newValue?.ToString();
                            Log($"已成功更新第{e.RowIndex + 1}行的处理结果字段: {oldValue} -> {newValue}");
                            isSpecialColumn = true;
                            break;
                        case "涉及企业":
                            oldValue = data.HeatingArea;
                            data.HeatingArea = newValue?.ToString();
                            Log($"已成功更新第{e.RowIndex + 1}行的涉及企业字段: {oldValue} -> {newValue}");
                            isSpecialColumn = true;
                            break;
                    }

                    // 处理普通列（有DataPropertyName的列）
                    if (!isSpecialColumn && property != null)
                    {
                        oldValue = property.GetValue(data);

                        // 更新对应的数据属性
                        if (!string.IsNullOrEmpty(columnName) && newValue != null)
                        {
                            if (property.CanWrite)
                            {
                                property.SetValue(data, Convert.ChangeType(newValue, property.PropertyType));
                                Log($"已成功更新第{e.RowIndex + 1}行的{columnName}字段: {oldValue} -> {newValue}");
                            }
                        }
                    }

                    // 如果编辑的是投诉内容列，重新匹配供热区域
                    if (isComplaintContentColumn && newValue is string newComplaintContent)
                    {
                        UpdateHeatingAreaForComplaint(data, newComplaintContent, e.RowIndex, originalHeatingArea);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"单元格编辑保存失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 根据投诉内容更新供热区域
        /// </summary>
        /// <param name="data">投诉数据对象</param>
        /// <param name="complaintContent">投诉内容</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="originalHeatingArea">原始供热区域</param>
        private void UpdateHeatingAreaForComplaint(ComplaintData data, string complaintContent, int rowIndex, string? originalHeatingArea)
        {
            try
            {
                if (_databaseService == null)
                {
                    Log("数据库服务未初始化，跳过供热区域匹配");
                    return;
                }

                Log($"开始为第{rowIndex + 1}行重新匹配供热区域");
                string? matchedHeatingArea = _databaseService.GetHeatingAreaByName(complaintContent);

                if (!string.IsNullOrEmpty(matchedHeatingArea))
                {
                    data.HeatingArea = matchedHeatingArea;
                    Log($"第{rowIndex + 1}行供热区域已更新: {originalHeatingArea} -> {matchedHeatingArea}");

                    // 刷新DataGridView以显示更新后的值
                    dataGridView.Refresh();
                }
                else
                {
                    Log($"第{rowIndex + 1}行未匹配到供热区域");
                }
            }
            catch (Exception ex)
            {
                Log($"重新匹配供热区域时出错: {ex.Message}");
            }
        }

        private void dataGridView_SelectionChanged(object sender, EventArgs e)
        {
            // 确保选择单元格后可以编辑
            if (sender is DataGridView dgv && dgv.CurrentCell != null && !dgv.ReadOnly)
            {
                dgv.ReadOnly = false;
                dgv.CurrentCell.ReadOnly = false;
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            _ = BeginInvoke(new Action(() =>
            {
                // 强制刷新整个控件层次
                SuspendLayout();

                // 确保DataGridView正确显示
                dataGridView.SuspendLayout();
                dataGridView.ColumnHeadersHeight = 80;
                dataGridView.ColumnHeadersVisible = true;
                dataGridView.Padding = new Padding(0, 10, 0, 0);  // 再次确认顶部内边距

                // 强制布局更新
                dataGridView.ResumeLayout(false);
                dataGridView.PerformLayout();

                // 更新整个窗体
                ResumeLayout(false);
                PerformLayout();

                // 强制重绘DataGridView
                dataGridView.Invalidate();
                dataGridView.Refresh();
            }));
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            // 确保窗口句柄已经创建
            if (IsHandleCreated)
            {
                _ = BeginInvoke(new Action(() =>
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
            }
        }
    }
}
