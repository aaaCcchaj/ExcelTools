using System;

namespace FS.Tool
{
    /// <summary>
    /// Excel坐标及其对应的属性名称
    /// </summary>
    public class ExcelRangeKeyDatasInfo : LinqObjectSerializableInfo
    {
        public ExcelRangeKeyDatasInfo(ILinqXmlSerializable xmlObject)
            : base(xmlObject)
        {
        }


        protected override void Initilize()
        {
        }

        /// <summary>
        /// 根据名称加载文件
        /// </summary>
        /// <param name="fileName"></param>
        public void LoadFile(string fileName)
        {
            LoadDocumentByFile(fileName);
        }
    }
}