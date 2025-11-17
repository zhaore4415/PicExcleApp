using System.Collections.Generic;

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
        /// 供热质量类关键词
        /// </summary>
        public string[] HeatingQualityKeywords { get; private set; }
        
        /// <summary>
        /// 维修类关键词
        /// </summary>
        public string[] MaintenanceKeywords { get; private set; }
        
        /// <summary>
        /// 政策咨询类关键词
        /// </summary>
        public string[] PolicyKeywords { get; private set; }
        
        /// <summary>
        /// 服务类关键词
        /// </summary>
        public string[] ServiceKeywords { get; private set; }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public KeywordConfig()
        {
            KeywordToCategoryMap = new Dictionary<string, string>()
            {
                { "不热", "质量问题" },
                { "维修", "维修问题" },
                { "漏水", "维修问题" },
                { "故障", "维修问题" }
            };
            
            CommunityToAreaMap = new Dictionary<string, string>();
            
            // 初始化公共关键词数组
            HeatingQualityKeywords = new string[] { "不热", "冰凉", "18", "温乎" };
            MaintenanceKeywords = new string[] { "管道", "地暖", "漏水", "爆管" };
            PolicyKeywords = new string[] { "退费", "收费", "缴费", "标准" };
            ServiceKeywords = new string[] { "态度", "不接", "没人接", "骂人" };
        }
    }
}