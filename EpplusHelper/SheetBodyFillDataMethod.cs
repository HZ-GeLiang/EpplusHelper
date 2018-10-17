namespace EpplusExtensions
{
    /// <summary>
    /// SheetBody数据填充方式
    /// </summary>
    public class SheetBodyFillDataMethod
    {
        public SheetBodyFillDataMethodOption FillDataMethodOption = default(SheetBodyFillDataMethodOption);

        public SynchronizationDataSourceConfig SynchronizationDataSource = new SynchronizationDataSourceConfig();

    }

    public enum SheetBodyFillDataMethodOption
    {
        /// <summary>
        /// 按约定填充: 单元格上配置了什么就填充什么
        /// </summary>
        Default = 0,
        /// <summary>
        /// 在约定填充的基础上,数据源Datatable的列如果没有被填充使用,那么将自动填充
        /// </summary>
        SynchronizationDataSource = 1

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
    }


}
