using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FS.Tool
{
    /// <summary>
    /// 读取ExcelRangeDatas数据信息
    /// </summary>
   public  interface IExcelRangeDatasOperator
    {
       /// <summary>
       /// 读取Excel的RangeData信息
       /// </summary>
       /// <param name="datas"></param>
        void ReadDataInfoFromRangeData(ExcelRangeKeyDatas datas);

        /// <summary>
        /// 获取写入Excel的RangeData信息
        /// </summary>
        void InitExcelRangeWriteData(ExcelRangeKeyDatas datas);
    }
}
