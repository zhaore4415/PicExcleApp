using Tesseract;
using System.IO;
using System;

namespace PicExcleApp.Services
{
    /// <summary>
    /// OCR文字识别服务
    /// </summary>
    public class OCRService
    {
        private readonly string _tessDataPath;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public OCRService()
        {
            // 简化路径设置，直接使用项目目录下的tessdata文件夹
            // 由于已设置文件较新复制，运行时会自动将文件复制到bin目录
            _tessDataPath = Path.Combine(Directory.GetCurrentDirectory(), "tessdata");
        }
        
        /// <summary>
        /// 从图片流中识别文字
        /// </summary>
        /// <param name="imageStream">图片流</param>
        /// <returns>识别出的文字</returns>
        public string RecognizeText(MemoryStream imageStream)
        {
            try
            {
                // 检查tessdata文件夹是否存在
                if (!Directory.Exists(_tessDataPath))
                {
                    Directory.CreateDirectory(_tessDataPath);
                    return "错误：未找到Tesseract语言包。请下载chi_sim.traineddata文件并放在tessdata文件夹中。";
                }
                
                // 检查是否有语言包文件
                string[] files = Directory.GetFiles(_tessDataPath, "*.traineddata");
                if (files.Length == 0)
                {
                    return "错误：tessdata文件夹为空。请从https://github.com/tesseract-ocr/tessdata下载chi_sim.traineddata和eng.traineddata文件并放入tessdata文件夹。";
                }
                
                // 检查是否有中文语言包
                if (!File.Exists(Path.Combine(_tessDataPath, "chi_sim.traineddata")))
                {
                    return "错误：缺少中文语言包。请下载chi_sim.traineddata文件并放入tessdata文件夹。";
                }
                
                // 检查是否有英文语言包
                if (!File.Exists(Path.Combine(_tessDataPath, "eng.traineddata")))
                {
                    return "错误：缺少英文语言包。请下载eng.traineddata文件并放入tessdata文件夹。";
                }
                
                using (var engine = new TesseractEngine(_tessDataPath, "chi_sim+eng", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromMemory(imageStream.ToArray()))
                    {
                        using (var page = engine.Process(img))
                        {
                            return page.GetText();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                // 提供更详细的错误信息
                if (ex.Message.Contains("Failed to initialise tesseract engine"))
                {
                    return "OCR识别失败：Tesseract引擎初始化失败。请确保已正确下载并放置语言包文件到tessdata文件夹。";
                }
                return $"OCR识别失败：{ex.Message}";
            }
        }
        
        /// <summary>
        /// 验证电话号码格式
        /// </summary>
        /// <param name="phone">电话号码</param>
        /// <returns>是否有效</returns>
        public bool ValidatePhone(string? phone)
        {
            // 简单的电话号码验证
            if (string.IsNullOrEmpty(phone))
                return false;
            return (System.Text.RegularExpressions.Regex.IsMatch(phone, @"^1[3-9]\d{9}$") ||
                    System.Text.RegularExpressions.Regex.IsMatch(phone, @"^\d{3,4}-\d{7,8}$"));
        }
        
        /// <summary>
        /// 验证工单号格式
        /// </summary>
        /// <param name="workOrderNumber">工单号</param>
        /// <returns>是否有效</returns>
        public bool ValidateWorkOrderNumber(string? workOrderNumber)
        {
            // 工单号可以是数字或字母数字组合
            if (string.IsNullOrEmpty(workOrderNumber))
                return false;
            return System.Text.RegularExpressions.Regex.IsMatch(workOrderNumber, @"^[A-Za-z0-9]+$");
        }
    }
}