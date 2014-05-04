using System;
using System.Xml.Linq;

namespace FS.Tool
{
    public class RangeData : IValidateItem, ILinqXmlSerializable
    {
        private const string ValidateStr = "{0}:不能为空！";

        #region .....字段.....
        /// <summary>
        /// 区域起始的单元格
        /// </summary>
        private string mFrom;

        /// <summary>
        /// 区域结束的单元格
        /// </summary>
        private string mTo;

        /// <summary>
        /// 对应的属性名称
        /// </summary>
        private string mPropertyName;

        /// <summary>
        /// 所在的Sheet页的名称
        /// </summary>
        private string mSheetName;

        /// <summary>
        /// 单元格的值
        /// </summary>
        private string mValue;

        /// <summary>
        /// 是否可以为空
        /// </summary>
        private bool mNotNull;

        /// <summary>
        /// 描述
        /// </summary>
        private string mDescription;
        #endregion .....字段.....

        #region .....属性.....
        /// <summary>
        /// 区域起始的单元格
        /// </summary>
        public string From
        {
            get
            {
                return this.mFrom;
            }
            set
            {
                if (mFrom != value)
                {
                    this.mFrom = value;
                }
            }
        }

        /// <summary>
        /// 区域结束的单元格
        /// </summary>
        public string To
        {
            get
            {
                return this.mTo;
            }
            set
            {
                if (mTo != value)
                {
                    this.mTo = value;
                }
            }
        }

        /// <summary>
        /// 对应的属性名称
        /// </summary>
        public string PropertyName
        {
            get
            {
                return this.mPropertyName;
            }
            set
            {
                if (mPropertyName != value)
                {
                    this.mPropertyName = value;
                }
            }
        }

        /// <summary>
        /// 所在的Sheet页的名称
        /// </summary>
        public string SheetName
        {
            get
            {
                return this.mSheetName;
            }
            set
            {
                if (mSheetName != value)
                {
                    this.mSheetName = value;
                }
            }
        }

        /// <summary>
        /// 单元格的值
        /// </summary>
        public string Value
        {
            get
            {
                return this.mValue;
            }
            set
            {
                if (mValue != value)
                {
                    this.mValue = value;
                }
            }
        }

        #endregion .....属性.....

        public void SetNotNull()
        {
            mNotNull = true;
        }

        public void SetNull()
        {
            mNotNull = false;
        }

        /// <summary>
        /// 验证通过
        /// </summary>
        public bool ValidatePass
        {
            get
            {
                if(mNotNull)
                {
                    return !string.IsNullOrEmpty(mValue);
                }
                return true;
            }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="sheetName"></param>
        public RangeData(string sheetName)
        {
            mSheetName = sheetName;
        }

        /// <summary>
        /// 验证的字符串
        /// </summary>
        public string ValidateString
        {
            get
            {
                if (!ValidatePass)
                {
                    return string.Format(ValidateStr, mDescription);
                }
                return string.Empty;
            }
        }

        public void WriteXml(XElement node)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 加载信息
        /// </summary>
        /// <param name="node"></param>
        public void ReadXml(XElement node)
        {
            mFrom = node.ReadString("From", string.Empty);
            mTo = node.ReadString("To", string.Empty);
            mPropertyName = node.Read("PropName", string.Empty);
            mNotNull = node.ReadBoolean("NotNull", false);
            mDescription = node.ReadString("Description", string.Empty);
        }

        /// <summary>
        /// 克隆
        /// </summary>
        /// <returns></returns>
        public RangeData Clone()
        {
            RangeData targetData = new RangeData(mSheetName);
            targetData.mFrom =mFrom;
            targetData.mTo = mTo;
            targetData.mPropertyName = mPropertyName;
            targetData.mValue = mValue;
            targetData.mNotNull = mNotNull;
            targetData.mDescription = mDescription;
            return targetData;
        }
    }
}