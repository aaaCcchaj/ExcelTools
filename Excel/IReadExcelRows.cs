using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FS.Tool
{
    /// <summary>
    /// 读取Excel信息实现的借口
    /// </summary>
    public interface IReadExcelRows
    {
        /// <summary>
        /// 需要读取的列数
        /// </summary>
        int ReadColumnNunbers { get; }

        /// <summary>
        /// 不能为空的列的索引
        /// </summary>
        int UnNullableColumnIndex { get; }

        /// <summary>
        /// 保存读取的每一行的值
        /// </summary>
        /// <param name="allColumnValue"></param>
        void AddReadColumnValues(string [] allColumnValue);

        /// <summary>
        /// 获取写入Excel的值得顺序集合
        /// </summary>
        /// <returns></returns>
        IEnumerable<string> ObjectValues();
    }
}
