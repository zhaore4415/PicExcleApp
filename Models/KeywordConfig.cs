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
        }
    }
}