using System;

namespace EPPlusExtensions
{
    /// <summary>
    /// SheetBody数据填充方式
    /// </summary>
    public class SheetBodyFillDataMethod
    {
        /// <summary>
        /// 填充数据选项
        /// </summary>
        public SheetBodyFillDataMethodOption FillDataMethodOption = SheetBodyFillDataMethodOption.Default;

        /// <summary>
        /// 填充数据同步方式
        /// </summary>
        public SynchronizationDataSourceConfig SynchronizationDataSource = new SynchronizationDataSourceConfig();

    }

    public enum SheetBodyFillDataMethodOption
    {
        /// <summary>
        /// 按约定填充: 单元格上配置了什么就填充什么
        /// </summary>
        Default = 1,
        /// <summary>
        /// 在约定填充的基础上,数据源Datatable的列如果没有被填充使用,那么将自动填充
        /// </summary>
        [Obsolete("不推荐使用,很多测试没通过")]
        SynchronizationDataSource = 2

    }

    public class SynchronizationDataSourceConfig
    {
        /// <summary>
        /// 需要同步Body
        /// </summary>
        public bool NeedBody { get; set; } = false;
        /// <summary>
        /// 需要同步Title
        /// </summary>
        public bool NeedTitle { get; set; } = true;
        /// <summary>
        /// 多余列只需要哪些列
        /// </summary>
        public string Include { get; set; }
        /// <summary>
        /// 多余列中不要那些列
        /// </summary>
        public string Exclude { get; set; }
    }


    public class FillDataColums
    {
        public FillDataColumsState State { get; set; }
        public string ColumName { get; set; }
    }

    public enum FillDataColumsState
    {
        /// <summary>
        /// 未使用
        /// </summary>
        Unchanged = 1,
        /// <summary>
        /// 已使用
        /// </summary>
        Used = 2,
        /// <summary>
        /// 将要使用
        /// </summary>
        WillUse = 3,
        /// <summary>
        /// 将不会使用
        /// </summary>
        WillNotUse = 4
    }

}
