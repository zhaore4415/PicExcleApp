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
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="ocrService">OCR服务</param>
        /// <param name="keywordConfig">关键词配置</param>
        public ContentParserService(OCRService ocrService, KeywordConfig keywordConfig)
        {
            _ocrService = ocrService;
            _keywordConfig = keywordConfig;
        }
        
        /// <summary>
        /// 解析OCR文本为投诉数据
        /// </summary>
        /// <param name="ocrText">OCR识别的文本</param>
        /// <param name="imagePath">图片路径</param>
        /// <returns>投诉数据</returns>
        public ComplaintData ParseToComplaintData(string ocrText, string imagePath)
        {
            var data = new ComplaintData
            {
                OriginalImagePath = imagePath,
                Status = "待处理"
            };
            
            try
            {
                // 提取工单号
                data.WorkOrderNumber = ExtractWorkOrderNumber(ocrText);
                
                // 提取姓名
                data.Name = ExtractName(ocrText);
                
                // 提取电话
                data.Phone = ExtractPhone(ocrText);
                
                // 提取投诉内容
                data.Content = ExtractContent(ocrText);
                
                // 提取来件日期
                data.CreateTime = ExtractDate(ocrText);
                
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
            var match = Regex.Match(text, @"DH\d{16}", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Value;
            
            // 工单号可能包含数字和字母
            match = Regex.Match(text, @"工单编[号:：]\s*([A-Za-z0-9]+)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;
            
            // 尝试通过关键词匹配，忽略字符间的空格
            match = Regex.Match(text, @"工\s*单\s*(编号|号|瘘|史)?\s*[：:}\s]*(DH\d{16})", RegexOptions.IgnoreCase);
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
            
            return "";
        }
        
        private string ExtractDate(string text)
        {
            // 预处理文本，移除多余空格但保留可能是日期一部分的空格
            var preprocessedText = text;
            
            // 尝试提取包含时间的标准日期格式
            var match = Regex.Match(preprocessedText, @"(\d{4}[-/]\d{1,2}[-/]\d{1,2}\s*\d{1,2}:\d{2})");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 尝试提取带年月日的中文日期格式
            match = Regex.Match(preprocessedText, @"(\d{4}年\d{1,2}月\d{1,2}日)");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 特别处理用户提供的特殊信访日期格式：2025¢10H26H 13:38
            match = Regex.Match(preprocessedText, @"(\d{4})[¢年]\s*(\d{1,2})[H月]\s*(\d{1,2})[H日]\s*(\d{1,2}):(\d{2})");
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 尝试提取带年月日时分的完整中文日期格式（重点匹配信访日期）
            match = Regex.Match(preprocessedText, @"(\d{4})\s*[年|H]\s*(\d{1,2})\s*[月|M]\s*(\d{1,2})\s*[日|H]\s*(\d{1,2})[：:;](\d{2})");
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 尝试处理OCR识别可能出错的日期格式，特别是带空格分隔的日期
            match = Regex.Match(preprocessedText, @"(\d{4})\s*[\sH¢年-/]\s*(\d{1,2})\s*[\sM月-/]\s*(\d{1,2})\s*[\sH日-/]\s*(\d{1,2})\s*[：:;]\s*(\d{2})");
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value.PadLeft(2, '0')}-{match.Groups[3].Value.PadLeft(2, '0')} {match.Groups[4].Value.PadLeft(2, '0')}:{match.Groups[5].Value}";
            }
            
            // 检查是否包含"提交"关键词附近的日期
            match = Regex.Match(preprocessedText, @"([提提交]\s*[交]?)\s*[：:;]?\s*(\d{4})[¢年H]\s*(\d{1,2})[H月M]\s*(\d{1,2})[H日]\s*(\d{1,2})[：:](\d{2})", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return $"{match.Groups[2].Value}-{match.Groups[3].Value.PadLeft(2, '0')}-{match.Groups[4].Value.PadLeft(2, '0')} {match.Groups[5].Value.PadLeft(2, '0')}:{match.Groups[6].Value}";
            }
            
            // 尝试匹配其他常见日期格式
            match = Regex.Match(text, @"(\d{2}[-/]\d{1,2}[-/]\d{1,2})");
            if (match.Success)
                return match.Groups[1].Value;
            
            return "";
        }
        
        private string ExtractContent(string text)
        {
            // 尝试提取投诉内容部分
            var content = text.Trim();
            
            // 移除已经提取的信息
            if (!string.IsNullOrEmpty(ExtractWorkOrderNumber(text)))
                content = Regex.Replace(content, $@"工单[号:：]?\s*{ExtractWorkOrderNumber(text)}", "", RegexOptions.IgnoreCase);
            
            if (!string.IsNullOrEmpty(ExtractName(text)))
                content = Regex.Replace(content, $@"姓名[：:]?\s*{ExtractName(text)}", "", RegexOptions.IgnoreCase);
            
            if (!string.IsNullOrEmpty(ExtractPhone(text)))
                content = Regex.Replace(content, $@"电话[：:]?\s*{ExtractPhone(text)}", "", RegexOptions.IgnoreCase);
            
            return content.Trim();
        }
        
        private string CategorizeComplaint(string content)
        {
            foreach (var keyword in _keywordConfig.KeywordToCategoryMap.Keys)
            {
                if (content.Contains(keyword))
                {
                    return _keywordConfig.KeywordToCategoryMap[keyword];
                }
            }
            
            return "正常问题";
        }
        
        private string MatchHeatingArea(string content)
        {
            foreach (var community in _keywordConfig.CommunityToAreaMap.Keys)
            {
                if (content.Contains(community))
                {
                    return _keywordConfig.CommunityToAreaMap[community];
                }
            }
            
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