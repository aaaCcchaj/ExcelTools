using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace FS.Tool
{
    public class OleDbExcel:IDisposable
    {
        public bool Success { get; set; }

        ///// <summary>
        ///// 连接字符串(2007的加载方式)
        ///// </summary>
        //private const string ConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties='Excel 8.0;HDR=no;IMEX=0'";

        /// <summary>
        /// 连接字符串（2003的加载方式）
        /// </summary>
        private const string ImportConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}; Extended Properties='Excel 8.0;HDR=no;IMEX=1'";
        private const string ExportConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}; Extended Properties='Excel 8.0;HDR=no;IMEX=0'";

        /// <summary>
        /// 查询指定Range的数据
        /// </summary>
        const string QueryDataFromRangeData = @"select * from [{0}$]";

        /// <summary>
        /// 连接
        /// </summary>
        private OleDbConnection mExcelConnection;

        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool OpenImportExcelFile(string fileName)
        {
            if(!File.Exists(fileName))
            {
                return false;
            }
            try
            {
                string excelConnString = string.Format(ImportConnString, fileName);
            mExcelConnection = new OleDbConnection(excelConnString);
            mExcelConnection.Open();
            return true;
            }
            catch (Exception eee)
            {
                Console.WriteLine(eee.Message);
                MessageBox.Show("打开文件失败，请查看文件是否已经打开，如果打开，请关闭之后重试！");
                Success = false;
                return false;
            }

        }
        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool OpenExportExcelFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return false;
            }
            try
            {
                string excelConnString = string.Format(ExportConnString, fileName);
                mExcelConnection = new OleDbConnection(excelConnString);
                mExcelConnection.Open();
                return true;
            }
            catch (Exception eee)
            {
                Success = false;
                Console.WriteLine(eee.Message);
                return false;
            }

        }

        public OleDbExcel(string fileName,bool import)
        {
            Success = true;
            if(import)
            {
                OpenImportExcelFile(fileName);
            }
            else
            {
                OpenExportExcelFile(fileName);
            }
        }

        /// <summary>
        /// 读取所有的信息
        /// </summary>
        /// <param name="datas"></param>
        public void GetRangesValue(ExcelRangeKeyDatas datas)
        {
            mCurrentSheetName = datas.SheetName;

            DataTable currentTable = new DataTable();

            var selectStr = string.Format(QueryDataFromRangeData, mCurrentSheetName);
            OleDbCommand olecommand = new OleDbCommand(selectStr, mExcelConnection);
            try
            {
                var dataReader = olecommand.ExecuteReader();
                currentTable.Load(dataReader);
                foreach (var rangeData in datas)
                {
                    var columnCharIndex = rangeData.From.FirstOrDefault();
                    var columnIndex = columnCharIndex - 'A';
                    Console.WriteLine(rangeData.From);
                    var row = Convert.ToInt32(rangeData.From.Remove(0, 1)) - 1;
                    rangeData.Value = currentTable.Rows[row][columnIndex].ToString();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(string.Format("================={0}================",e.Message));
                return;
            }

            foreach (var excelRangeKeyData in datas)
            {

            }
        }

        private string mCurrentSheetName;


        /// <summary>
        /// 读取RangeData
        /// </summary>
        public string GetRangeValue(RangeData rangeData)
        {
            var sheetName = rangeData.SheetName;
            mCurrentSheetName = rangeData.SheetName;
            var selectStr = string.Format(QueryDataFromRangeData, sheetName);
            OleDbCommand olecommand = new OleDbCommand(selectStr, mExcelConnection);
            try
            {
                var reader = olecommand.ExecuteReader();
                DataTable table  =new DataTable();
                table.Load(reader);
                var columnCharIndex = rangeData.From.FirstOrDefault();
                var columnIndex = columnCharIndex - 'A';
                var row = Convert.ToInt32(rangeData.From.Remove(0, 1)) - 1;
                rangeData.Value = table.Rows[row][columnIndex].ToString();
                return rangeData.Value;

            }
            catch (Exception e)
            {
                return string.Empty;
            }
            return string.Empty;

        }

        /// <summary>
        /// 读取RangeData
        /// </summary>
        public void SetRangeValue(RangeData rangeData)
        {
            InitilizeExcelTool(rangeData.SheetName);
            SetRangeValueCore(rangeData);
        }



        DataTable mSetDataTable;
        Dictionary<char, DataColumn> mColumnIndex;
        OleDbDataAdapter mDataAdapter;
        DataSet mDS;

        /// <summary>
        /// 设置RangeData
        /// </summary>
        private string SetRangeValueCore(RangeData rangeData)
        {
            try
            {
                char columnCharIndex = rangeData.From.FirstOrDefault();
                int rowIndex = int.Parse(rangeData.From.Remove(0, 1));
                string upDataCommand = "update [{0}$] set {1}='{2}' where {3}={4}";
                var oleDbCommandText = string.Format(upDataCommand,rangeData.SheetName,mColumnIndex[columnCharIndex].ColumnName,rangeData.Value,mColumnIndex.Last().Value.ColumnName,rowIndex);
                OleDbCommand command = new OleDbCommand(oleDbCommandText,mExcelConnection);
                command.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                return string.Empty;
            }
            return string.Empty;

        }

        /// <summary>
        /// 读取RangeData
        /// </summary>
        public string SetRangeValues(ExcelRangeKeyDatas rangeDatas)
        {

            try
            {
                InitilizeExcelTool(rangeDatas.SheetName);
                foreach (RangeData rangedata in rangeDatas)
                {
                    SetRangeValueCore(rangedata);
                }

            }
            catch (Exception e)
            {
                return string.Empty;
            }
            return string.Empty;

        }

        /// <summary>
        /// 初始化Excel工具
        /// </summary>
        public void InitilizeExcelTool(string tableName)
        {
            var sheetName = tableName;
            mCurrentSheetName = tableName;
            var selectStr = string.Format(QueryDataFromRangeData, sheetName);
            mDS = new DataSet();
            mDataAdapter = new OleDbDataAdapter(selectStr, mExcelConnection);
            mDataAdapter.Fill(mDS);
            if (mDS.Tables.Count == 0)
            {
            }
            mSetDataTable = mDS.Tables[0];
            InitializeColumnIndex();
        }


        /// <summary>
        /// 初始化列的索引
        /// </summary>
        private void InitializeColumnIndex()
        {
            if(mSetDataTable != null)
            {
                mColumnIndex = new Dictionary<char, DataColumn>();
                int columnIndex = (int)'A';
                foreach (DataColumn column in mSetDataTable.Columns)
                {
                    mColumnIndex.Add((char)columnIndex, column);
                    columnIndex++;
                }
            }
            //if (mColumnIndex.Count > 0)
            //{
            //    StringBuilder upDataCommandStr = new StringBuilder();
            //    upDataCommandStr.Append("update [");
            //    upDataCommandStr.Append(tableName);
            //    upDataCommandStr.Append("$] set ");
            //    foreach (var column in mColumnIndex)
            //    {
            //        upDataCommandStr.Append(string.Format("{0}=@{1}，",column.Value.ColumnName,column.Key));
            //    }
            //    upDataCommandStr.Remove(upDataCommandStr.Length - 1, 1);
            //    upDataCommandStr.Append(" ");
            //    upDataCommandStr.Append(string.Format("where {0}=@{1}", mColumnIndex.Last().Value.ColumnName, mColumnIndex.Last().Key));
            //    upSetCmd = new OleDbCommand(upDataCommandStr.ToString(), mExcelConnection);
            //    mDataAdapter.UpdateCommand = upSetCmd;

            //}
        }


        /// <summary>
        /// 析构
        /// </summary>
        public void Dispose()
        {
            mExcelConnection.Close();
        }
    }
}
