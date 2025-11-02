using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using PicExcleApp.Models;

namespace PicExcleApp.Services
{
    /// <summary>
    /// Excel导出服务
    /// </summary>
    public class ExcelExportService
    {
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
                
                // 自动调整列宽
                worksheet.Cells.AutoFitColumns();
                
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
    }
}