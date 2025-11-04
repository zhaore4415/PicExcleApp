using System.Data.SQLite;
using System.Data;
using System.IO;
using System.Collections.Generic;

namespace PicExcleApp.Services
{
    /// <summary>
    /// 数据库服务类，用于处理SQLite数据库操作
    /// </summary>
    public class DatabaseService
    {
        private readonly string _dbPath;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="dbPath">数据库文件路径</param>
        public DatabaseService(string dbPath = null)
        {
            // 如果没有提供路径，使用项目根目录下的heating.db
            _dbPath = string.IsNullOrEmpty(dbPath) ? Path.Combine(Directory.GetCurrentDirectory(), "heating.db") : dbPath;
            
        }
        
        /// <summary>
        /// 根据投诉内容模糊匹配供热区域
        /// 首先尝试匹配ResidentialName，然后尝试匹配BuildingName
        /// </summary>
        /// <param name="complaintContent">投诉内容</param>
        /// <returns>匹配到的供热区域名称，如果未匹配到则返回"未知"</returns>
        public string GetHeatingAreaByName(string complaintContent)
        {
            if (string.IsNullOrEmpty(complaintContent))
            {
                Log("投诉内容为空，返回未知区域");
                return "未知";
            }
            
            // 预处理投诉内容，提高匹配成功率
            string processedContent = PreprocessText(complaintContent);
            Log($"开始根据投诉内容查找供热区域（预处理后）: {processedContent}");
            
            string heatingAreaName = "未知";
            
            // 1. 首先尝试匹配小区名称(ResidentialName)
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
                {
                    connection.Open();
                    
                    // 先获取所有小区信息，然后在内存中进行更智能的匹配
                    string query = @"
                    SELECT r.ResidentialName, ha.HeatingAreaName 
                    FROM t_Residential r
                    JOIN t_HeatingArea ha ON r.HeatingAreaId = ha.Id
                    GROUP BY r.ResidentialName, ha.HeatingAreaName  -- 去重
                    ORDER BY LENGTH(r.ResidentialName) DESC;  -- 优先匹配较长的名称
                    ";
                    
                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string residentialName = reader["ResidentialName"].ToString() ?? string.Empty;
                            string areaName = reader["HeatingAreaName"].ToString() ?? string.Empty;
                            
                            if (!string.IsNullOrEmpty(residentialName) && !string.IsNullOrEmpty(areaName))
                            {
                                // 尝试多种匹配方式
                                string processedResName = residentialName;
                                
                                // 1. 完全包含匹配
                                if (complaintContent.Contains(residentialName) || processedContent.Contains(processedResName))
                                {
                                    heatingAreaName = areaName;
                                    Log($"通过小区名称完全匹配成功: '{residentialName}' -> '{heatingAreaName}'");
                                    return heatingAreaName;
                                }
                                
                                // 2. 相似度匹配（适用于OCR识别错误或部分匹配）
                                if (GetSimilarityScore(processedContent, processedResName) > 0.6)
                                {
                                    heatingAreaName = areaName;
                                    Log($"通过小区名称相似度匹配成功: '{residentialName}' -> '{heatingAreaName}'");
                                    return heatingAreaName;
                                }
                            }
                        }
                    }
                }
                Log("小区名称匹配失败，尝试楼栋名称匹配");
            }
            catch (System.Exception ex)
            {
                Log($"小区名称查询异常: {ex.Message}");
            }
            
            // 2. 如果小区名称未匹配到，尝试匹配楼栋名称(BuildingName)
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
                {
                    connection.Open();
                    
                    // 先获取所有楼栋信息，然后在内存中进行更智能的匹配
                    string query = @"
                    SELECT r.BuildingName as BuildingName, ha.HeatingAreaName  as HeatingAreaName
                    FROM t_Residential r
                    
                    JOIN t_HeatingArea ha ON r.HeatingAreaId = ha.Id
                    where  ha.Id=24
                    
                    ;  
                    ";
                    
                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string buildingName = reader[0].ToString() ?? string.Empty;
                            string areaName = reader[1].ToString() ?? string.Empty;

                            if (!string.IsNullOrEmpty(buildingName) && !string.IsNullOrEmpty(areaName) && buildingName.Length >= 2)
                            {
                                // 尝试多种匹配方式
                                string processedBldName = buildingName;
                                
                                // 1. 完全包含匹配
                                if (complaintContent.Contains(buildingName) || processedContent.Contains(processedBldName))
                                {
                                    heatingAreaName = areaName;
                                    Log($"通过楼栋名称完全匹配成功: '{buildingName}' -> '{heatingAreaName}'");
                                    return heatingAreaName;
                                }
                                
                                // 2. 相似度匹配（楼栋名称需要更高的相似度阈值）
                                if (GetSimilarityScore(processedContent, processedBldName) > 0.7)
                                {
                                    heatingAreaName = areaName;
                                    Log($"通过楼栋名称相似度匹配成功: '{buildingName}' -> '{heatingAreaName}'");
                                    return heatingAreaName;
                                }
                            }
                        }
                    }
                }
                Log("楼栋名称匹配失败");
            }
            catch (System.Exception ex)
            {
                Log($"楼栋名称查询异常: {ex.Message}");
            }
            
            return heatingAreaName;
        }
        
        /// <summary>
        /// 预处理文本，用于更好的模糊匹配
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <returns>处理后的文本</returns>
        private string PreprocessText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;
                
            // 移除多余空格、制表符和换行符
            string processed = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();
            // 移除常见的特殊字符，但保留中文字符和数字
            processed = System.Text.RegularExpressions.Regex.Replace(processed, @"[^\u4e00-\u9fa5\d\s]", "");
            // 转为小写进行不区分大小写的匹配
            return processed.ToLower();
        }
        
        /// <summary>
        /// 计算两个字符串的相似度（基于Jaccard系数的简化版本）
        /// </summary>
        /// <param name="s1">第一个字符串</param>
        /// <param name="s2">第二个字符串</param>
        /// <returns>相似度分数（0-1之间）</returns>
        private double GetSimilarityScore(string s1, string s2)
        {
            try
            {
                if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
                    return 0;
                
                // 使用字符集计算交集和并集
                HashSet<char> set1 = new HashSet<char>(s1);
                HashSet<char> set2 = new HashSet<char>(s2);
                
                // 计算交集
                int intersection = set1.Intersect(set2).Count();
                if (intersection == 0)
                    return 0;
                
                // 计算并集
                int union = set1.Union(set2).Count();
                
                // 计算Jaccard系数
                return (double)intersection / union;
            }
            catch
            {
                return 0;
            }
        }
        
        /// <summary>
        /// 记录日志信息
        /// </summary>
        /// <param name="message">日志消息</param>
        private void Log(string message)
        {
            try
            {
                string logPath = Path.Combine(Directory.GetCurrentDirectory(), "database_service.log");
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";
                File.AppendAllText(logPath, logMessage);
            }
            catch
            {
                // 日志记录失败不影响程序运行
            }
        }
        
        /// <summary>
        /// 从数据库中加载所有小区到供热区域的映射（用于兼容现有代码）
        /// </summary>
        /// <returns>小区名称到供热区域的映射字典</returns>
        public Dictionary<string, string> GetAllCommunityToAreaMap()
        {
            // 使用不区分大小写的字典，提高匹配成功率
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
                {
                    connection.Open();
                    
                    // 查询所有小区和楼栋信息
                    string query = @"
                    SELECT r.ResidentialName, r.BuildingName, ha.HeatingAreaName 
                    FROM t_Residential r
                    JOIN t_HeatingArea ha ON r.HeatingAreaId = ha.Id
                    GROUP BY r.ResidentialName, r.BuildingName, ha.HeatingAreaName;
                    ";
                    
                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    {
                        using (SQLiteDataReader reader = command.ExecuteReader())
                        {
                            // 存储唯一的小区-区域映射
                            var residentialMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            
                            while (reader.Read())
                            {
                                string residentialName = reader["ResidentialName"].ToString() ?? string.Empty;
                                string buildingName = reader["BuildingName"].ToString() ?? string.Empty;
                                string heatingAreaName = reader["HeatingAreaName"].ToString() ?? string.Empty;
                                
                                if (!string.IsNullOrEmpty(residentialName) && !string.IsNullOrEmpty(heatingAreaName))
                                {
                                    // 为小区名称创建多种变体形式
                                    AddToMapWithVariations(map, residentialName, heatingAreaName);
                                    
                                    // 保存唯一的小区-区域映射用于后续处理变体
                                    if (!residentialMap.ContainsKey(residentialName))
                                    {
                                        residentialMap[residentialName] = heatingAreaName;
                                    }
                                    
                                    // 为楼栋名称也创建映射（当楼栋名称较长时）
                                    if (!string.IsNullOrEmpty(buildingName) && buildingName.Length >= 2)
                                    {
                                        AddToMapWithVariations(map, buildingName, heatingAreaName);
                                        
                                        // 创建小区+楼栋组合的映射
                                        string combinedName = $"{residentialName}{buildingName}";
                                        AddToMapWithVariations(map, combinedName, heatingAreaName);
                                    }
                                }
                            }
                            
                            // 添加一些常见的小区名称变体形式
                            AddCommonVariations(map, residentialMap);
                        }
                    }
                }
                
                Log($"成功加载 {map.Count} 个小区/楼栋到供热区域的映射关系");
            }
            catch (System.Exception ex)
            {
                Log($"加载小区到供热区域映射关系时发生错误: {ex.Message}");
            }
            
            return map;
        }
        
        /// <summary>
        /// 为一个名称添加多种变体形式到映射字典中
        /// </summary>
        /// <param name="map">映射字典</param>
        /// <param name="name">原始名称</param>
        /// <param name="value">供热区域名称</param>
        private void AddToMapWithVariations(Dictionary<string, string> map, string name, string value)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(value))
                return;
            
            // 原始名称
            if (!map.ContainsKey(name))
                map[name] = value;
            
            // 预处理后的名称
            string processedName = PreprocessText(name);
            if (!string.IsNullOrEmpty(processedName) && !map.ContainsKey(processedName))
                map[processedName] = value;
            
            // 小写形式
            string lowerName = name.ToLower();
            if (!map.ContainsKey(lowerName))
                map[lowerName] = value;
            
            // 移除空格的形式
            string noSpaceName = name.Replace(" ", "");
            if (!string.IsNullOrEmpty(noSpaceName) && !map.ContainsKey(noSpaceName))
                map[noSpaceName] = value;
        }
        
        /// <summary>
        /// 添加一些常见的小区名称变体形式
        /// </summary>
        /// <param name="map">完整的映射字典</param>
        /// <param name="residentialMap">小区-区域的基础映射</param>
        private void AddCommonVariations(Dictionary<string, string> map, Dictionary<string, string> residentialMap)
        {
            // 常见的小区名称后缀和前缀变体
            string[] commonSuffixes = { "小区", "花园", "家园", "苑", "庄", "园", "公寓", "大厦", "楼", "栋", "座" };
            
            foreach (var kvp in residentialMap)
            {
                string residentialName = kvp.Key;
                string heatingAreaName = kvp.Value;
                
                // 尝试移除常见后缀
                foreach (string suffix in commonSuffixes)
                {
                    if (residentialName.EndsWith(suffix))
                    {
                        string nameWithoutSuffix = residentialName.Substring(0, residentialName.Length - suffix.Length).Trim();
                        if (!string.IsNullOrEmpty(nameWithoutSuffix))
                        {
                            AddToMapWithVariations(map, nameWithoutSuffix, heatingAreaName);
                        }
                    }
                }
                
                // 尝试添加常见后缀
                foreach (string suffix in commonSuffixes)
                {
                    if (!residentialName.EndsWith(suffix) && residentialName.Length > 2)
                    {
                        string nameWithSuffix = $"{residentialName}{suffix}";
                        if (!map.ContainsKey(nameWithSuffix))
                        {
                            map[nameWithSuffix] = heatingAreaName;
                        }
                    }
                }
            }
        }
    }
}