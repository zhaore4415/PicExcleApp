using Newtonsoft.Json;
using System.IO;
using PicExcleApp.Models;

namespace PicExcleApp.Services
{
    /// <summary>
    /// 配置管理服务
    /// </summary>
    public class ConfigManagerService
    {
        private readonly string _configFilePath;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public ConfigManagerService()
        {
            _configFilePath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");
        }
        
        /// <summary>
        /// 加载配置
        /// </summary>
        /// <returns>关键词配置</returns>
        public KeywordConfig LoadConfig()
        {   
            if (File.Exists(_configFilePath))
            {   
                try
                {   
                    string json = File.ReadAllText(_configFilePath);
                    return JsonConvert.DeserializeObject<KeywordConfig>(json) ?? new KeywordConfig();
                }
                catch
                {   
                    return new KeywordConfig();
                }
            }
            
            // 返回默认配置
            return new KeywordConfig();
        }
        
        /// <summary>
        /// 保存配置
        /// </summary>
        /// <param name="config">关键词配置</param>
        public void SaveConfig(KeywordConfig config)
        {   
            try
            {   
                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(_configFilePath, json);
            }
            catch (System.Exception ex)
            {   
                throw new System.Exception($"保存配置失败：{ex.Message}");
            }
        }
    }
}