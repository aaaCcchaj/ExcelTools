using System;

namespace FS.Tool
{
    /// <summary>
    /// ExcelRangeDatasExtend的扩展方法
    /// </summary>
    public static class ExcelRangeDatasExtend
    {
        /// <summary>
        /// 读取时间类型
        /// </summary>
        /// <param name="datas"></param>
        /// <param name="dataIndex"></param>
        /// <returns></returns>
        public static DateTime ReadDateTime(this ExcelRangeKeyDatas datas, string dataIndex)
        {
            var dateTime = DateTime.MinValue;
            var convertString = datas[dataIndex];
            if (string.IsNullOrEmpty(convertString))
            {
                return dateTime;
            }
            if (string.IsNullOrEmpty(convertString))
            {
                return dateTime;
            }
            DateTime.TryParse(convertString, out dateTime);
            return dateTime;
        }

        /// <summary>
        /// 读取bool类型
        /// </summary>
        /// <param name="datas"></param>
        /// <param name="dataIndex"></param>
        /// <returns></returns>
        public static bool ReadBoolean(this ExcelRangeKeyDatas datas, string dataIndex)
        {
            bool boolValue;
            var convertString = datas[dataIndex];
            if (string.IsNullOrEmpty(convertString))
            {
                return false;
            }
            Boolean.TryParse(convertString, out boolValue);
            return boolValue;
        }

        /// <summary>
        /// 读取bool类型
        /// </summary>
        /// <param name="datas"></param>
        /// <param name="dataIndex"></param>
        /// <returns></returns>
        public static short ReadShort(this ExcelRangeKeyDatas datas, string dataIndex)
        {
            short shortValue;
            var convertString = datas[dataIndex];
            if (string.IsNullOrEmpty(convertString))
            {
                return 0;
            }
            Int16.TryParse(convertString, out shortValue);
            return shortValue;
        }
    }
}
