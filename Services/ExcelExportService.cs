using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using PicExcleApp.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace PicExcleApp.Services
{
    /// <summary>
    /// Excel导出服务
    /// </summary>
    public class ExcelExportService
    {
        private readonly KeywordConfig _keywordConfig;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="keywordConfig">关键词配置</param>
        public ExcelExportService(KeywordConfig keywordConfig = null)
        {
            _keywordConfig = keywordConfig ?? new KeywordConfig();
        }
        /// <summary>
        /// 导出数据到Excel文件
        /// </summary>
        /// <param name="dataList">投诉数据列表</param>
        /// <param name="filePath">导出文件路径</param>
        public void ExportToExcel(List<ComplaintData> dataList, string filePath)
        {
            // 设置EPPlus许可上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("德惠市城区供热企业信访统计表");
                
                // 第1行：标题行
                worksheet.Cells["A1:J1"].Merge = true;
                worksheet.Cells["A1"].Value = "德惠市城区供热企业信访统计表";
                worksheet.Cells["A1"].Style.Font.Size = 14;
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                
                // 第2行：表头行
                worksheet.Cells["A2"].Value = "序号";
                worksheet.Cells["B2"].Value = "信访日期";
                worksheet.Cells["C2"].Value = "工单号";
                worksheet.Cells["D2"].Value = "信访来源";
                worksheet.Cells["E2"].Value = "分类";
                worksheet.Cells["F2"].Value = "涉及企业";
                worksheet.Cells["G2"].Value = "投诉内容";
                worksheet.Cells["H2"].Value = "联系电话";
                worksheet.Cells["I2"].Value = "测温温度";
                worksheet.Cells["J2"].Value = "处理结果";
                
                // 设置表头样式（黄色背景）
                using (var range = worksheet.Cells["A2:J2"])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    // 添加边框
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                
                // 从第3行开始填充数据
                for (int i = 0; i < dataList.Count; i++)
                {
                    var data = dataList[i];
                    var row = i + 3; // 从第3行开始填充数据
                    
                    worksheet.Cells[row, 1].Value = i + 1; // 序号从1开始
                    worksheet.Cells[row, 2].Value = data.CreateTime ?? string.Empty; // 信访日期
                    worksheet.Cells[row, 3].Value = data.WorkOrderNumber ?? string.Empty; // 工单号
                    worksheet.Cells[row, 4].Value = data.Source ?? string.Empty; // 信访来源
                    worksheet.Cells[row, 5].Value = data.Category ?? string.Empty; // 分类
                    worksheet.Cells[row, 6].Value = data.HeatingArea ?? "正德"; // 涉及企业（默认为正德）
                    worksheet.Cells[row, 7].Value = data.Content ?? string.Empty; // 投诉内容
                    worksheet.Cells[row, 8].Value = data.Phone ?? string.Empty; // 联系电话
                    worksheet.Cells[row, 9].Value = data.Temperature ?? string.Empty; // 测温温度
                    worksheet.Cells[row, 10].Value = data.Result ?? string.Empty; // 处理结果
                    
                    // 为数据行添加边框
                    using (var range = worksheet.Cells[row, 1, row, 10])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }
                }
                
                // 手动设置各列宽度
                // 这里可以根据需要调整每列的宽度值
                worksheet.Column(1).Width = 8;  // 序号列
                worksheet.Column(2).Width = 15; // 信访日期列
                worksheet.Column(3).Width = 15; // 工单号列
                worksheet.Column(4).Width = 12; // 信访来源列
                worksheet.Column(5).Width = 10; // 分类列
                worksheet.Column(6).Width = 12; // 涉及企业列
                worksheet.Column(7).Width = 40; // 投诉内容列（可以设置较宽）
                worksheet.Column(8).Width = 15; // 联系电话列
                worksheet.Column(9).Width = 10; // 测温温度列
                worksheet.Column(10).Width = 20; // 处理结果列
                
                // 如果需要自动调整列宽而不是手动设置，可以取消上面的注释并使用下面这行
                // worksheet.Cells.AutoFitColumns();
                
                // 保存文件
                File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
        }
        
        /// <summary>
        /// 导入小区映射文件
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>小区到区域的映射字典</returns>
        public Dictionary<string, string> ImportCommunityMap(string filePath)
        {
            var map = new Dictionary<string, string>();
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                
                // 假设第一列是小区名称，第二列是供热区域
                for (int row = 2; row <= rowCount; row++) // 从第二行开始，第一行是表头
                {
                    string community = worksheet.Cells[row, 1].Text;
                    string area = worksheet.Cells[row, 2].Text;
                    
                    if (!string.IsNullOrEmpty(community) && !string.IsNullOrEmpty(area))
                    {
                        map[community] = area;
                    }
                }
            }
            
            return map;
        }
        
        /// <summary>
        /// 导入关键词配置文件
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>关键词到分类的映射字典</returns>
        public Dictionary<string, string> ImportKeywordMap(string filePath)
        {
            var map = new Dictionary<string, string>();
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                
                // 假设第一列是关键词，第二列是分类
                for (int row = 2; row <= rowCount; row++) // 从第二行开始，第一行是表头
                {
                    string keyword = worksheet.Cells[row, 1].Text;
                    string category = worksheet.Cells[row, 2].Text;
                    
                    if (!string.IsNullOrEmpty(keyword) && !string.IsNullOrEmpty(category))
                    {
                        map[keyword] = category;
                    }
                }
            }
            
            return map;
        }
        /// <summary>
        /// 从data.xls文件提取企业信息
        /// </summary>
        /// <param name="rootDirectory">根目录路径</param>
        /// <summary>
        /// 导入燃气表数据
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>投诉数据列表</returns>
        public List<ComplaintData> ImportGasMeterData(string filePath)
        {
            var dataList = new List<ComplaintData>();
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                
                // 从第二行开始读取数据
                for (int row = 2; row <= rowCount; row++)
                {
                    var complaintData = new ComplaintData();

                    // 尝试从不同列获取数据（根据实际Excel结构调整）
                    complaintData.WorkOrderNumber = worksheet.Cells[row, 1]?.Text; // 工单号
                    complaintData.Content = worksheet.Cells[row, 4]?.Text; // 投诉内容
                    complaintData.Phone = worksheet.Cells[row, 13]?.Text; // 联系电话
                    complaintData.HeatingArea = "";
                    complaintData.Result = "";
                    complaintData.Temperature = "";
                    complaintData.Source = "燃气表数据";
                    complaintData.CreateTime = worksheet.Cells[row, 8]?.Text; // 信仿日期交办时间

                    // 基于关键词的分类赋值逻辑
                    string content = complaintData.Content ?? string.Empty;
                    if (!string.IsNullOrEmpty(content))
                    {
                        // 使用公共关键词配置
                        // 检查是否包含供热质量类关键词
                        bool isHeatingQuality = _keywordConfig.HeatingQualityKeywords.Any(keyword => content.Contains(keyword));
                        // 检查是否包含维修类关键词
                        bool isMaintenance = _keywordConfig.MaintenanceKeywords.Any(keyword => content.Contains(keyword));
                        // 检查是否包含政策咨询类关键词
                        bool isPolicy = _keywordConfig.PolicyKeywords.Any(keyword => content.Contains(keyword));
                        // 检查是否包含服务类关键词
                        bool isService = _keywordConfig.ServiceKeywords.Any(keyword => content.Contains(keyword));

                        // 设置分类优先级：供热质量类 > 维修类 > 政策咨询类 > 服务类
                        if (isHeatingQuality)
                            complaintData.Category = "供热质量类";
                        else if (isMaintenance)
                            complaintData.Category = "维修类";
                        else if (isPolicy)
                            complaintData.Category = "政策咨询类";
                        else if (isService)
                            complaintData.Category = "服务类";
                        else
                            complaintData.Category = "无";
                    }
                    else
                    {
                        complaintData.Category = "无";
                    }

                    dataList.Add(complaintData);
                }
            }
            
            return dataList;
        }
        
        /// <summary>
        /// 从 data.xls 文件提取企业信息
        /// </summary>
        /// <param name="rootDirectory">根目录路径</param>
        /// <param name="complaintContents">投诉内容列表</param>
        /// <returns>投诉内容到企业名称的映射字典</returns>
        public Dictionary<string, string> ExtractEnterpriseInfoFromDataFile(string rootDirectory, List<string> complaintContents)
        {
            var resultMap = new Dictionary<string, string>();
            string dataFilePath = Path.Combine(rootDirectory, "data.xls");

            // 检查文件是否存在
            if (!File.Exists(dataFilePath))
            {
                Console.WriteLine($"文件不存在: {dataFilePath}");
                return resultMap; // 返回空字典
            }

            Console.WriteLine($"开始处理 data.xls 文件: {dataFilePath}");
            Console.WriteLine($"投诉内容列表数量: {complaintContents.Count}");

            try
            {
                // 使用 FileStream 打开文件
                using (var fileStream = new FileStream(dataFilePath, FileMode.Open, FileAccess.Read))
                {
                    // 创建 HSSFWorkbook 对象来处理.xls 文件
                    HSSFWorkbook workbook = new HSSFWorkbook(fileStream);

                    Console.WriteLine($"成功打开 Excel 文件，工作表数量: {workbook.NumberOfSheets}");

                    // 第一阶段：遍历所有工作表的B列进行匹配
                    Console.WriteLine("===== 第一阶段：遍历所有工作表的B列进行匹配 =====");
                    foreach (var content in complaintContents)
                    {
                        if (resultMap.ContainsKey(content))
                            continue;

                        for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                        {
                            ISheet worksheet = workbook.GetSheetAt(sheetIndex);
                            if (worksheet.LastRowNum < 0) continue;

                            // 提取企业名称
                            string enterpriseName = ExtractEnterpriseNameFromNPOI(worksheet);
                            if (string.IsNullOrEmpty(enterpriseName))
                                enterpriseName = worksheet.SheetName;

                            // 收集并匹配B列关键词
                            List<string> columnBKeywords = new List<string>();
                            for (int row = 2; row <= Math.Min(worksheet.LastRowNum, 399); row++)
                            {
                                IRow currentRow = worksheet.GetRow(row);
                                if (currentRow != null)
                                {
                                    ICell bCell = currentRow.GetCell(1);
                                    string bCellValue = GetCellValue(bCell)?.Trim();
                                    if (!string.IsNullOrEmpty(bCellValue) && !bCellValue.Equals("小区名称", StringComparison.OrdinalIgnoreCase))
                                    {
                                        columnBKeywords.Add(bCellValue);
                                    }
                                }
                            }

                            // 尝试严格匹配
                            bool matched = MatchKeywords(content, columnBKeywords, enterpriseName, resultMap, $"工作表 '{worksheet.SheetName}' 的B列");
                            if (matched)
                                break;

                            // 尝试模糊匹配
                            matched = FuzzyMatchKeywords(content, columnBKeywords, enterpriseName, resultMap, $"工作表 '{worksheet.SheetName}' 的B列");
                            if (matched)
                                break;
                        }
                    }

                    // 统计第一阶段匹配结果
                    Console.WriteLine($"第一阶段（B列匹配）完成，已匹配: {resultMap.Count} 条记录");

                    // 第二阶段：对未匹配的内容，遍历所有工作表的C列进行匹配
                    Console.WriteLine("===== 第二阶段：遍历所有工作表的C列进行匹配 =====");
                    foreach (var content in complaintContents)
                    {
                        if (resultMap.ContainsKey(content))
                            continue;

                        for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                        {
                            ISheet worksheet = workbook.GetSheetAt(sheetIndex);
                            if (worksheet.LastRowNum < 0) continue;

                            // 提取企业名称
                            string enterpriseName = ExtractEnterpriseNameFromNPOI(worksheet);
                            if (string.IsNullOrEmpty(enterpriseName))
                                enterpriseName = worksheet.SheetName;

                            // 收集并匹配C列关键词
                            List<string> columnCKeywords = new List<string>();
                            for (int row = 2; row <= Math.Min(worksheet.LastRowNum, 99); row++)
                            {
                                IRow currentRow = worksheet.GetRow(row);
                                if (currentRow != null)
                                {
                                    ICell cCell = currentRow.GetCell(2);
                                    string cCellValue = GetCellValue(cCell)?.Trim();
                                    if (!string.IsNullOrEmpty(cCellValue) && !cCellValue.Equals("楼房名称", StringComparison.OrdinalIgnoreCase))
                                    {
                                        columnCKeywords.Add(cCellValue);
                                    }
                                }
                            }

                            // 尝试严格匹配
                            bool matched = MatchKeywords(content, columnCKeywords, enterpriseName, resultMap, $"工作表 '{worksheet.SheetName}' 的C列");
                            if (matched)
                                break;

                            // 尝试模糊匹配
                            matched = FuzzyMatchKeywords(content, columnCKeywords, enterpriseName, resultMap, $"工作表 '{worksheet.SheetName}' 的C列");
                            if (matched)
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取 data.xls 文件时出错: {ex.Message}");
                Console.WriteLine($"错误类型: {ex.GetType().Name}");
                Console.WriteLine($"错误详情: {ex.StackTrace}");
            }

            Console.WriteLine($"企业信息提取完成，总共匹配到 {resultMap.Count} 条记录");
            return resultMap;
        }
        
        /// <summary>
        /// 匹配关键词 - 使用更严格的匹配逻辑
        /// </summary>
        private bool MatchKeywords(string content, List<string> keywords, string enterpriseName, Dictionary<string, string> resultMap, string columnName)
        {
            foreach (var keyword in keywords)
            {
                if (!string.IsNullOrEmpty(keyword) && keyword.Length > 1) // 至少2个字符才匹配
                {
                    // 使用更精确的匹配，确保关键词是完整的小区名称等
                    if (content.Contains(keyword))
                    {
                        // 对于包含企业名称关键词的情况，需要更谨慎处理
                        if (enterpriseName.Contains("万兴") && !content.Contains("万兴花园"))
                        {
                            // 如果企业名称包含"万兴"但投诉内容不包含"万兴花园"，跳过匹配
                            continue;
                        }
                        
                        resultMap[content] = enterpriseName;
                        Console.WriteLine($"{columnName}匹配成功: '{keyword}' 在 '{content.Substring(0, Math.Min(50, content.Length))}...' -> 企业: {enterpriseName}");
                        return true;
                    }
                }
            }
            return false;
        }
        
        /// <summary>
        /// 模糊匹配关键词 - 提高模糊匹配的准确性，避免过度匹配
        /// </summary>
        private bool FuzzyMatchKeywords(string content, List<string> keywords, string enterpriseName, Dictionary<string, string> resultMap, string columnName)
        {
            foreach (var keyword in keywords)
            {
                if (!string.IsNullOrEmpty(keyword) && keyword.Length > 3) // 模糊匹配要求至少4个字符
                {
                    // 尝试关键词的部分匹配
                    string[] parts = keyword.Split(' ', '\t', '，', ',', '\n', '\r');
                    foreach (var part in parts)
                    {
                        // 部分匹配要求至少3个字符，减少误匹配
                        if (part.Length > 2 && content.Contains(part))
                        {
                            // 对于可能导致误匹配的情况进行额外检查
                            if (enterpriseName.Contains("万兴") && !content.Contains("万兴花园"))
                            {
                                // 如果企业名称包含"万兴"但投诉内容不包含"万兴花园"，跳过模糊匹配
                                continue;
                            }
                            
                            resultMap[content] = enterpriseName;
                            Console.WriteLine($"{columnName}模糊匹配成功: '{part}' (来自 '{keyword}') 在 '{content.Substring(0, Math.Min(50, content.Length))}...' -> 企业: {enterpriseName}");
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 从NPOI工作表中提取企业名称（支持从"供热站名称：恒泰"格式中提取"恒泰"）
        /// </summary>
        /// <param name="worksheet">NPOI Excel工作表</param>
        /// <returns>提取的企业名称</returns>
        private string ExtractEnterpriseNameFromNPOI(ISheet worksheet)
        {
            // 尝试多个可能的位置提取企业名称
            // 定义行和列的索引组合（行索引从0开始，列索引从0开始）
            // (行, 列) 对应 Excel 单元格位置
            // 优先从A2和B1单元格提取
            var potentialLocations = new[]
            {
                (row: 1, col: 0), // A2
                (row: 0, col: 1), // B1
                (row: 1, col: 1), // B2
                (row: 0, col: 0), // A1
                (row: 1, col: 2)  // C2
            };
            
            foreach (var (row, col) in potentialLocations)
            {
                try
                {
                    IRow sheetRow = worksheet.GetRow(row);
                    if (sheetRow != null)
                    {
                        ICell cell = sheetRow.GetCell(col);
                        string cellValue = GetCellValue(cell)?.Trim();
                        
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            char[] separators = { '：', ':' };
                            Console.WriteLine($"检查单元格 ({row+1},{col+1}): '{cellValue}'");
                            
                            // 检查是否包含供热站名称标识
                            if (cellValue.Contains("供热站名称") || cellValue.Contains("供热站"))
                            {
                                // 处理"供热站名称：恒泰"格式
                                foreach (char separator in separators)
                                {
                                    if (cellValue.Contains(separator))
                                    {
                                        string name = cellValue.Split(separator)[1].Trim();
                                        if (!string.IsNullOrEmpty(name)) 
                                        {
                                            Console.WriteLine($"从格式'供热站名称{separator}XXX'中提取企业名称: {name}");
                                            return name;
                                        }
                                    }
                                }
                            }
                            
                            // 清理可能的前缀并返回
                            if (cellValue.Contains("站") || cellValue.Contains("公司") || cellValue.Contains("供热"))
                            {
                                // 清理常见前缀
                                if (cellValue.StartsWith("供热站名称：", StringComparison.OrdinalIgnoreCase))
                                    return cellValue.Substring(5).Trim();
                                if (cellValue.StartsWith("供热站名称:", StringComparison.OrdinalIgnoreCase))
                                    return cellValue.Substring(5).Trim();
                                return cellValue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"提取单元格 ({row+1},{col+1}) 时出错: {ex.Message}");
                }
            }
            
            return null;
        }

        private string GetCellValue(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString();
                    }
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    IFormulaEvaluator evaluator = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                    CellValue evaluatedValue = evaluator.Evaluate(cell);
                    return GetCellValueFromEvaluatedCell(evaluatedValue);
                default:
                    return null;
            }
        }

        private string GetCellValueFromEvaluatedCell(CellValue cellValue)
        {
            if (cellValue == null)
            {
                return null;
            }

            switch (cellValue.CellType)
            {
                case CellType.String:
                    return cellValue.StringValue;
                case CellType.Numeric:
                    // 对于CellValue类型，不使用DateUtil.IsCellDateFormatted，直接处理数值
                    return cellValue.NumberValue.ToString();
                case CellType.Boolean:
                    return cellValue.BooleanValue.ToString();
                default:
                    return null;
            }
        }

        

    }
}