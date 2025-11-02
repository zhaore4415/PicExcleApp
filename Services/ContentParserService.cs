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
            // 工单号可能包含数字和字母
            var match = Regex.Match(text, @"工单[号:：]\s*([A-Za-z0-9]+)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;
            
            // 尝试直接匹配数字序列
            match = Regex.Match(text, @"\b\d{6,12}\b");
            if (match.Success)
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
            // 提取日期格式：2025-10-13 12:38 或 2025/10/13 12:38 或 2025年10月13日
            var match = Regex.Match(text, @"(\d{4}[-/]\d{1,2}[-/]\d{1,2}\s*\d{1,2}:\d{2})");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 尝试匹配单独的日期格式
            match = Regex.Match(text, @"(\d{4}年\d{1,2}月\d{1,2}日)");
            if (match.Success)
                return match.Groups[1].Value;
            
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