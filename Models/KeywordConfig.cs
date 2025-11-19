using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace PicExcleApp.Models
{
    /// <summary>
    /// 关键词配置类
    /// </summary>
    public class KeywordConfig
    {
        /// <summary>
        /// 关键词与分类的映射
        /// </summary>
        public Dictionary<string, string> KeywordToCategoryMap { get; set; }
        
        /// <summary>
        /// 小区名称与供热区域的映射
        /// </summary>
        public Dictionary<string, string> CommunityToAreaMap { get; set; }
        
        
        /// <summary>
        /// 类别名称到关键词列表的映射
        /// </summary>
        public Dictionary<string, List<string>> CategoryToKeywordsMap { get; private set; }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public KeywordConfig()
        {
            KeywordToCategoryMap = new Dictionary<string, string>();
            CommunityToAreaMap = new Dictionary<string, string>();
            CategoryToKeywordsMap = new Dictionary<string, List<string>>();
        
            // 从Excel文件加载类别和关键词
            LoadCategoriesFromExcel();
        }
        
        /// <summary>
        /// 从Excel文件加载类别和关键词
        /// </summary>
        private void LoadCategoriesFromExcel()
        {
            // 使用相对路径，优先在当前应用目录查找class.xlsx
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string excelFilePath = Path.Combine(appDirectory, "class.xlsx");
            
            // 如果相对路径不存在，尝试使用原始的绝对路径作为备选
            if (!File.Exists(excelFilePath))
            {
                string backupPath = @"D:\南禾网站建设\苏州跨麦\PicExcle\PicExcleApp\class.xlsx";
                if (File.Exists(backupPath))
                {
                    excelFilePath = backupPath;
                }
            }
            
            // 检查文件是否存在
            if (!File.Exists(excelFilePath))
            {
                // 如果文件不存在，使用默认关键词初始化映射
                return;
            }
            
            try
            {
                using (FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    
                    // 根据文件扩展名创建相应的工作簿对象
                    if (excelFilePath.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else if (excelFilePath.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        return;
                    }
                    
                    // 获取第一个工作表
                    ISheet sheet = workbook.GetSheetAt(0);
                    if (sheet == null)
                    {
                        return;
                    }
                    
                    // 清空现有映射
                    KeywordToCategoryMap.Clear();
                    CategoryToKeywordsMap.Clear();
                    
                    // 遍历每一行
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null) continue;
                        
                        // 获取类别名称（A列）
                    ICell categoryCell = row.GetCell(0);
                    if (categoryCell == null) continue;
                    
                    string categoryName = GetCellStringValue(categoryCell)?.Trim();
                    if (string.IsNullOrEmpty(categoryName)) continue;
                    
                    // 如果类别不存在于映射中，添加它
                    if (!CategoryToKeywordsMap.ContainsKey(categoryName))
                    {
                        CategoryToKeywordsMap[categoryName] = new List<string>();
                    }
                    
                    // 遍历B列及以后的所有列，获取关键词
                    for (int cellIndex = 1; cellIndex < row.LastCellNum; cellIndex++)
                    {
                        ICell keywordCell = row.GetCell(cellIndex);
                        if (keywordCell == null) continue;
                        
                        string keyword = GetCellStringValue(keywordCell)?.Trim();
                        if (!string.IsNullOrEmpty(keyword))
                        {
                            // 添加关键词到类别映射
                            CategoryToKeywordsMap[categoryName].Add(keyword);
                            // 添加关键词到分类映射
                            if (!KeywordToCategoryMap.ContainsKey(keyword))
                            {
                                KeywordToCategoryMap[keyword] = categoryName;
                            }
                        }
                    }
                    }
                 
                }
            }
            catch (Exception ex)
            {
               
            }
        }
     
        
        /// <summary>
        /// 重新加载Excel文件中的类别和关键词
        /// </summary>
        public void ReloadCategories()
        {
            LoadCategoriesFromExcel();
        }
        
        /// <summary>
        /// 获取单元格的字符串值，无论单元格类型是什么
        /// </summary>
        /// <param name="cell">Excel单元格</param>
        /// <returns>单元格的字符串值</returns>
        private string GetCellStringValue(ICell cell)
        {
            if (cell == null)
                return null;
                
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    // 对于数值类型，转换为字符串
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    try
                    {
                        // 尝试获取公式计算结果
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.String:
                                return cell.StringCellValue;
                            case CellType.Numeric:
                                return cell.NumericCellValue.ToString();
                            case CellType.Boolean:
                                return cell.BooleanCellValue.ToString();
                            default:
                                return null;
                        }
                    }
                    catch
                    {
                        return null;
                    }
                default:
                    return null;
            }
        }
    }
}