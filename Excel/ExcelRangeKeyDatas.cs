using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using System.Linq;

namespace FS.Tool
{
    /// <summary>
    /// Excel坐标及其对应的属性名称
    /// </summary>
    public class ExcelRangeKeyDatas : List<RangeData>, IValidateItem
    {
        #region .....字段.....
        /// <summary>
        /// sheet名字
        /// </summary>
        private string mSheetName;

        /// <summary>
        /// 读取模式
        /// </summary>
        private ExcelReadMode mMode;

        /// <summary>
        /// 关键列（行），不能为空的
        /// </summary>
        private string mKey;

        /// <summary>
        /// 开始行（列）
        /// </summary>
        private string mStart;

        #endregion .....字段.....

        #region .....属性.....
        /// <summary>
        /// 获取数据
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public string this[string propertyName]
        {
            get
            {
                var rangeData = GetDataByPropName(propertyName);
                if (rangeData != null && !string.IsNullOrEmpty(rangeData.Value))
                {
                    return rangeData.Value;
                }
                return string.Empty;
            }
        }

        public RangeData GetDataByPropName(string propertyName)
        {
            return this.FirstOrDefault(rData => rData.PropertyName == propertyName);
            ;
        }

        /// <summary>
        /// Excel的读取模式
        /// </summary>
        public ExcelReadMode ExcelMode
        {
            get { return mMode; }
        }

        /// <summary>
        /// 获取关键列（行）
        /// </summary>
        public string Key
        {
            get { return mKey; }
        }

        /// <summary>
        /// 开始列/行
        /// </summary>
        public string Start
        {
            get { return mStart; }
        }

        /// <summary>
        /// sheet名字
        /// </summary>
        public string SheetName
        {
            get
            {
                return this.mSheetName;
            }
            set
            {
                this.mSheetName = value;
            }
        }
        #endregion .....属性.....

        #region .....方法.....
        /// <summary>
        /// 加载
        /// </summary>
        /// <param name="filePath"></param>
        public void Load(string filePath)
        {
            ExcelRangeKeyDatasInfo loadInfo = new ExcelRangeKeyDatasInfo(this);
            loadInfo.LoadDocumentByFile(filePath);
            InitializeReadModel();
        }

        /// <summary>
        /// 设置指定的值
        /// </summary>
        /// <param name="propertyName"></param>
        /// <param name="value"></param>
        public void SetDataValue(string propertyName, string value)
        {
            var rangeData = this.FirstOrDefault(rData => rData.PropertyName == propertyName);
            if (rangeData != null)
            {
                rangeData.Value = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void InitializeReadModel()
        {
            switch (mMode)
            {
                case ExcelReadMode.Custom:
                    return;
                case ExcelReadMode.RowRecyle:
                    InitilizeRowRecyle();
                    return;
                case ExcelReadMode.ColumnRecyle:

                    return;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void InitilizeRowRecyle()
        {
             mStartRow = Convert.ToInt32(mStart);
         }

        /// <summary>
        /// 开始行
        /// </summary>
        private int mStartRow;

        /// <summary>
        /// 初始化列的循环信息
        /// </summary>
        public RangeData GetNewRowRecyle(ExcelRangeKeyDatas datas)
        {
            RangeData keyData = null;
            datas.SheetName = this.mSheetName;
            foreach (var rangeData in this)
            {
                RangeData data = new RangeData(mSheetName);
                data.PropertyName = rangeData.PropertyName;
                data.From = string.Format("{0}{1}", rangeData.From, mStartRow);
                var strrTo = string.IsNullOrEmpty(rangeData.To) ? rangeData.From : rangeData.To;
                data.To = string.Format("{0}{1}", strrTo, mStartRow);
                if (rangeData.From == Key)
                {
                    keyData = data;
                }
                datas.Add(data);
            }
            mStartRow++;
            return keyData;
        }

        /// <summary>
        /// 初始化列的循环信息
        /// </summary>
        public void InitializeRowRecyle(int rowIndex)
        {
            foreach (var rangeData in this)
            {
                rangeData.PropertyName = rangeData.PropertyName;
                rangeData.From = string.Format("{0}{1}", rangeData.From, rowIndex);
                rangeData.To = string.IsNullOrEmpty(rangeData.To) ? rangeData.From : string.Format("{0}{1}", rangeData.To, rowIndex); 
            }
        }

        /// <summary>
        /// 写入
        /// </summary>
        /// <param name="node"></param>
        public void WriteXml(XElement node)
        {
        }

        /// <summary>
        /// 读取
        /// </summary>
        /// <param name="node"></param>
        public void ReadXml(XElement node)
        {
            mSheetName = node.ReadString("SheetName", string.Empty);
            var modeStr = node.ReadString("Mode", string.Empty);
            if(string.IsNullOrEmpty(modeStr))
            {
                throw new Exception("Excel读取配置文件的读取模式不能为空");
            }
            mKey = node.ReadString("Key", string.Empty);
            mStart = node.ReadString("Start", string.Empty);
            Enum.TryParse(modeStr, out mMode);
            var rangeNodes = node.Elements("RangeData");
            foreach (var rangeNode in rangeNodes)
            {
                RangeData data = new RangeData(mSheetName);
                data.ReadXml(rangeNode);
                Add(data);
            }
        }
        #endregion .....方法.....

        /// <summary>
        /// 克隆
        /// </summary>
        /// <returns></returns>
        public ExcelRangeKeyDatas Clone()
        {
            ExcelRangeKeyDatas cloneTarget=new ExcelRangeKeyDatas();
            cloneTarget.mSheetName = mSheetName;
            cloneTarget.mStartRow = mStartRow;
            cloneTarget.mStart = mStart;
            cloneTarget.mMode = mMode;
            cloneTarget.mKey = mKey;
            foreach (var item in this)
            {
                cloneTarget.Add(item.Clone());
            }
            return cloneTarget;
        }

        /// <summary>
        /// 导入时，是否验证通过
        /// </summary>
        public bool ValidatePass
        {
            get
            {
                return this.All(item => (item.ValidatePass));
            }
        }

        /// <summary>
        /// 导入时验证的字符串
        /// </summary>
        public string ValidateString
        {
            get
            {
                if(ValidatePass)
                {
                    return string.Empty;
                }
                StringBuilder validateStringBuilder = new StringBuilder();
                validateStringBuilder.Append("============");
                validateStringBuilder.Append(this.mSheetName);
                validateStringBuilder.AppendLine("=========");

                foreach (var item in this.Where(item=>!(item.ValidatePass)))
                {
                    validateStringBuilder.AppendLine(item.ValidateString);
                }
                validateStringBuilder.AppendLine("============================");
                return validateStringBuilder.ToString();
            }
        }
    }

    /// <summary>
    /// excel文件的读取方式
    /// </summary>
    public enum ExcelReadMode
    {
        /// <summary>
        /// 行循环
        /// </summary>
        RowRecyle,

        /// <summary>
        /// 自定义
        /// </summary>
        Custom,

        /// <summary>
        /// 列循环
        /// </summary>
        ColumnRecyle
    }
}
