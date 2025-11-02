using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Formats.Bmp;
using System.IO;

namespace PicExcleApp.Services
{
    /// <summary>
    /// 图片处理服务
    /// </summary>
    public class ImageProcessingService
    {
        /// <summary>
        /// 预处理图片
        /// </summary>
        /// <param name="imagePath">图片路径</param>
        /// <returns>处理后的图片流</returns>
        public MemoryStream PreprocessImage(string imagePath)
        {
            using (var image = SixLabors.ImageSharp.Image.Load(imagePath))
            {
                // 调整方向
                image.Mutate(x => x.AutoOrient());
                
                // 增强对比度
                image.Mutate(x => x.Contrast(0.1f));
                
                // 转换为灰度图以提高OCR准确性
                image.Mutate(x => x.Grayscale());
                
                // 保存到内存流
                var stream = new MemoryStream();
                image.Save(stream, new PngEncoder());
                stream.Position = 0;
                return stream;
            }
        }
        
        /// <summary>
        /// 检查文件是否为支持的图片格式
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否为支持的图片格式</returns>
        public bool IsSupportedImageFormat(string filePath)
        {
            var extensions = new[] { ".jpg", ".jpeg", ".png", ".bmp" };
            var extension = Path.GetExtension(filePath).ToLower();
            return extensions.Contains(extension);
        }
    }
}