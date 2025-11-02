using System;

namespace PicExcleApp.Models
{
    /// <summary>
    /// 投诉数据模型
    /// </summary>
    public class ComplaintData
    {
        /// <summary>
        /// 工单号
        /// </summary>
        public string? WorkOrderNumber { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// 电话
        /// </summary>
        public string? Phone { get; set; }

        /// <summary>
        /// 投诉内容
        /// </summary>
        public string? Content { get; set; }

        /// <summary>
        /// 分类
        /// </summary>
        public string? Category { get; set; }

        /// <summary>
        /// 所属供热区域
        /// </summary>
        public string? HeatingArea { get; set; }

        /// <summary>
        /// 原始图片路径
        /// </summary>
        public string? OriginalImagePath { get; set; }
        
        /// <summary>
        /// 来件日期
        /// </summary>
        public string? CreateTime { get; set; }
        
        /// <summary>
        /// 处理状态
        /// </summary>
        public string? Status { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }
        
        /// <summary>
        /// 信访来源
        /// </summary>
        public string? Source { get; set; }
        
        /// <summary>
        /// 测温温度
        /// </summary>
        public string? Temperature { get; set; }
        
        /// <summary>
        /// 处理结果
        /// </summary>
        public string? Result { get; set; }
    }
}