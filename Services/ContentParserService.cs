using System.Text.RegularExpressions;
using PicExcleApp.Models;

namespace PicExcleApp.Services
{
    /// <summary>
    /// 内容解析服务
    /// </summary>
    public class ContentParserService
    {
        private readonly OCRService _ocrService;
        private readonly KeywordConfig _keywordConfig;
        private readonly DatabaseService _databaseService;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="ocrService">OCR服务</param>
        /// <param name="keywordConfig">关键词配置</param>
        public ContentParserService(OCRService ocrService, KeywordConfig keywordConfig)
        {
            _ocrService = ocrService;
            _keywordConfig = keywordConfig;
            _databaseService = new DatabaseService();
            
            // 加载数据库中的小区映射到KeywordConfig中，保持与现有代码的兼容性
            LoadCommunityMapFromDatabase();
        }
        
        /// <summary>
        /// 从数据库加载小区到供热区域的映射
        /// </summary>
        private void LoadCommunityMapFromDatabase()
        {
            try
            {
                var communityMap = _databaseService.GetAllCommunityToAreaMap();
                if (communityMap.Count > 0)
                {
                    // 清空现有的映射并添加数据库中的映射
                    _keywordConfig.CommunityToAreaMap.Clear();
                    foreach (var kvp in communityMap)
                    {
                        _keywordConfig.CommunityToAreaMap[kvp.Key] = kvp.Value;
                    }
                }
            }
            catch (System.Exception ex)
            {
                // 记录错误但不中断程序运行
                System.Console.WriteLine($"加载数据库小区映射失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 去除汉字之间的多余空格和行与行之间的多余空行
        /// </summary>
        /// <param name="text">需要处理的文本</param>
        /// <returns>处理后的文本</returns>
        private string RemoveExtraSpaces(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
                
            // 使用正则表达式移除汉字之间的所有空格
            // 匹配模式：一个汉字，后跟任意数量的空格，再跟一个汉字
            string result = text;
            
            // 重复应用正则替换，直到没有更多的空格被移除
            string prevResult;
            do
            {
                prevResult = result;
                // 移除两个汉字之间的空格
                result = Regex.Replace(result, @"([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])", "$1$2");
            } while (prevResult != result);
            
            // 处理行与行之间的多余空行，包括可能包含空白字符的空行
            result = Regex.Replace(result, @"\r?\n\s*\r?\n", "\r\n");
            
            // 重复处理直到没有更多的空行被移除
            string prevResult2;
            do
            {
                prevResult2 = result;
                result = Regex.Replace(result, @"\r?\n\s*\r?\n", "\r\n");
            } while (prevResult2 != result);
            
            return result;
        }
        
        /// <summary>
        /// 解析OCR文本为投诉数据
        /// </summary>
        /// <param name="ocrText">OCR识别的文本</param>
        /// <param name="imagePath">图片路径</param>
        /// <returns>投诉数据</returns>
        public ComplaintData ParseToComplaintData(string ocrText, string imagePath)
        {
            // 预处理OCR文本，移除汉字之间的多余空格
            string processedText = RemoveExtraSpaces(ocrText);
            
            var data = new ComplaintData
            {
                OriginalImagePath = imagePath,
                Status = "待处理"
            };
            
            try
            {
                // 提取工单号
                data.WorkOrderNumber = ExtractWorkOrderNumber(processedText);
                
                // 提取姓名
                data.Name = ExtractName(processedText);
                
                // 提取电话
                data.Phone = ExtractPhone(processedText);
                
                // 提取投诉内容
                data.Content = ExtractContent(processedText);
                
                // 提取来件日期
                data.CreateTime = ExtractDate(processedText);
                
                // 自动分类
                data.Category = CategorizeComplaint(data.Content);
                
                // 自动匹配供热区域
                data.HeatingArea = MatchHeatingArea(data.Content);
                
                // 验证关键字段
                ValidateData(data);
                
                data.Status = "成功";
            }
            catch (System.Exception ex)
            {
                data.Status = "失败";
                data.ErrorMessage = ex.Message;
            }
            
            return data;
        }
        
        private string ExtractWorkOrderNumber(string text)
        {
            // 查找以DH开头的工单号（用户明确指出工单号格式为DH开头）
            var match = Regex.Match(text, @"DH\d{17}", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Value;
            
            // 工单号可能包含数字和字母
            match = Regex.Match(text, @"工单编[号:：]\s*([A-Za-z0-9]+)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;
            
            // 尝试通过关键词匹配，忽略字符间的空格
            match = Regex.Match(text, @"工\s*单\s*(编号|号|瘘|史)?\s*[：:}\s]*(DH\d{17})", RegexOptions.IgnoreCase);
            if (match.Success && !string.IsNullOrEmpty(match.Groups[2].Value))
                return match.Groups[2].Value;
            
            // 尝试匹配较长的数字序列
            match = Regex.Match(text, @"\b[A-Za-z]?\d{8,20}\b", RegexOptions.IgnoreCase);
            if (match.Success && match.Value.Length >= 8)
                return match.Value;
            
            return "";
        }
        
        private string ExtractName(string text)
        {
            // 提取姓名
            var match = Regex.Match(text, @"姓名[：:]\s*([\u4e00-\u9fa5a-zA-Z]+)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;
            
            // 尝试匹配可能的姓名格式
            match = Regex.Match(text, @"[\u4e00-\u9fa5]{2,4}\s*");
            if (match.Success)
                return match.Value.Trim();
            
            return "";
        }
        
        private string ExtractPhone(string text)
        {
            // 提取电话号码
            var match = Regex.Match(text, @"1[3-9]\d{9}");
            if (match.Success)
                return match.Value;
            
            match = Regex.Match(text, @"\d{3,4}-\d{7,8}");
            if (match.Success)
                return match.Value;
            
            return "****"; // 当电话为空时显示****
        }
        
        private string ExtractDate(string text)
        {
            // 预处理文本，移除多余空格但保留可能是日期一部分的空格
            var preprocessedText = text;
            
            // 1. 优先处理用户指定的特殊格式：在"提交"关键词之前查找的日期格式
            // 匹配模式：任何字符序列，后跟日期格式(可能包含特殊字符)，然后是与"提交"相关的关键词
            var match = Regex.Match(preprocessedText, @"(\d{4})[¢年H]\s*(\d{1,2})[H月M]\s*(\d{1,2})[H日D]\s+(\d{1,2})[：:](\d{2})\s+[提提交]", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 2. 直接匹配用户提供的特殊信访日期格式：2025¢10H26H 13:38
            match = Regex.Match(preprocessedText, @"(\d{4})[¢年]\s*(\d{1,2})[H月]\s*(\d{1,2})[H日]\s+(\d{1,2}):(\d{2})");
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 3. 针对OCR识别错误的特殊格式优化，处理字符间可能存在的乱码或错误字符
            match = Regex.Match(preprocessedText, @"(\d{4})[¢¤]\s*(\d{1,2})[Hh]\s*(\d{1,2})[Hh]\s+(\d{1,2})[：:](\d{2})");
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 4. 检查是否包含"提交"关键词附近的日期（日期在提交后）
            match = Regex.Match(preprocessedText, @"[提提交]\s*[交]?\s*[于在]?\s*(\d{4})[¢年H]\s*(\d{1,2})[H月M]\s*(\d{1,2})[H日]\s*(\d{1,2})[：:](\d{2})", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 5. 检查是否包含"于"关键词后的日期
            match = Regex.Match(preprocessedText, @"[于在]\s*(\d{4})[¢年H]\s*(\d{1,2})[H月M]\s*(\d{1,2})[H日]\s*(\d{1,2})[：:](\d{2})", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 6. 尝试提取包含时间的标准日期格式
            match = Regex.Match(preprocessedText, @"(\d{4}[-/]\d{1,2}[-/]\d{1,2}\s*\d{1,2}:\d{2})");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 7. 尝试提取带年月日的中文日期格式
            match = Regex.Match(preprocessedText, @"(\d{4}年\d{1,2}月\d{1,2}日)");
            if (match.Success)
                return match.Groups[1].Value;
            
            return "";
        }
        
        private string ExtractContent(string text)
        {
            // 尝试提取"诉求内容"关键词后面的文字
            string content = string.Empty;
            
            // 使用正则表达式查找"诉求内容"关键词及其后面的所有文字
            var match = Regex.Match(text, @"诉求内容[:：]?\s*(.+)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (match.Success && match.Groups.Count > 1)
            {
                content = match.Groups[1].Value.Trim();
            }
            else
            {
                // 如果没有找到"诉求内容"关键词，则使用原始逻辑
                content = text.Trim();
                
                // 移除已经提取的信息
                if (!string.IsNullOrEmpty(ExtractWorkOrderNumber(text)))
                    content = Regex.Replace(content, $@"工单[号:：]?\s*{ExtractWorkOrderNumber(text)}", "", RegexOptions.IgnoreCase);
                
                if (!string.IsNullOrEmpty(ExtractName(text)))
                    content = Regex.Replace(content, $@"姓名[：:]?\s*{ExtractName(text)}", "", RegexOptions.IgnoreCase);
                
                // 只在电话号码确实存在（不是空字符串且不是"****"）时才移除
                var phoneNumber = ExtractPhone(text);
                if (!string.IsNullOrEmpty(phoneNumber) && phoneNumber != "****")
                    content = Regex.Replace(content, $@"电话[：:]?\s*{phoneNumber}", "", RegexOptions.IgnoreCase);
            }
            
            // 移除投诉内容中汉字之间的所有空格
            content = RemoveExtraSpaces(content);
            
            return content.Trim();
        }
        
        private string CategorizeComplaint(string content)
        {
            // 完全优先使用从Excel动态加载的关键词和类别映射
            foreach (var keyword in _keywordConfig.KeywordToCategoryMap.Keys)
            {
                if (content.Contains(keyword))
                {
                    return _keywordConfig.KeywordToCategoryMap[keyword];
                }
            }
            
            // 如果动态映射中没有匹配到，遍历所有动态加载的类别和关键词
            // 这样可以确保所有从Excel加载的类别（包括"测试类"）都能被正确匹配
            foreach (var category in _keywordConfig.CategoryToKeywordsMap.Keys)
            {
                foreach (var keyword in _keywordConfig.CategoryToKeywordsMap[category])
                {
                    if (content.Contains(keyword))
                    {
                        return category;
                    }
                }
            }
            
            return "无";
        }
        
        private string MatchHeatingArea(string content)
        {
            try
            {
                // 注释掉从heating.db查询供热站的代码，现在从data.xls获取企业信息
                // 优先使用数据库进行模糊匹配
                //string heatingArea = _databaseService.GetHeatingAreaByName(content);
                //if (!string.IsNullOrEmpty(heatingArea) && heatingArea != "未知")
                //{
                //    return heatingArea;
                //}
            }
            catch (System.Exception ex)
            {
                // 数据库查询失败时，回退到使用配置文件中的映射
                System.Console.WriteLine($"数据库查询供热区域失败: {ex.Message}");
            }
            
            // 如果数据库查询失败或未找到匹配项，回退到使用配置文件中的映射
            //foreach (var community in _keywordConfig.CommunityToAreaMap.Keys)
            //{
            //    if (content.Contains(community))
            //    {
            //        return _keywordConfig.CommunityToAreaMap[community];
            //    }
            //}
            
            return "未知";
        }
        
        private void ValidateData(ComplaintData data)
        {
            var errors = new System.Collections.Generic.List<string>();
            
            if (!_ocrService.ValidateWorkOrderNumber(data.WorkOrderNumber))
            {
                errors.Add("工单号格式不正确");
            }
            
            if (!_ocrService.ValidatePhone(data.Phone))
            {
                errors.Add("电话号码格式不正确");
            }
            
            if (errors.Count > 0)
            {
                data.ErrorMessage = string.Join("; ", errors);
            }
        }
    }
}