using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace FS.Tool
{
    public class ExcelTool : IDisposable
    {

        static readonly object MissingValue = Missing.Value;
        
        #region field
        private Microsoft.Office.Interop.Excel.Application mExcelClass;

        /// <summary>
        /// 工作文档
        /// </summary>
        private Microsoft.Office.Interop.Excel._Workbook mWorkBook;

        /// <summary>
        /// 提高效率用的缓存sheet
        /// </summary>
        private Microsoft.Office.Interop.Excel.Worksheet mTempSheet;
        #endregion


        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName"></param>
        public ExcelTool(string fileName)
        {
            Debug.Assert(!string.IsNullOrEmpty(fileName));
            Debug.Assert(File.Exists(fileName));
            mExcelClass = new Application { Visible = false };
            mExcelClass.DisplayAlerts = false;
            mWorkBook = mExcelClass.Workbooks.Open(fileName, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue,
                                    MissingValue, MissingValue, MissingValue, MissingValue);
        }

        /// <summary>
        /// 打开新的WorkBook
        /// </summary>
        /// <param name="fileName"></param>
        public void OpenNewWorkBook(string fileName)
        {
            if (mWorkBook != null)
            {
                mWorkBook.Close();
            }
            mWorkBook = mExcelClass.Workbooks.Open(fileName, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue,
                                    MissingValue, MissingValue, MissingValue, MissingValue);
        }

        #region .....获取/设定指定位置的值.....

        /// <summary>
        /// 获取指定sheet的指定格的值
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <returns></returns>
        private string GetRangeValue(string sheetName, string from, string to)
        {
            if (mWorkBook == null)
            {
                return string.Empty;
            }
            if (mTempSheet == null || mTempSheet.Name != sheetName)
            {
                mTempSheet = null;
                for (int i = 0; i < mWorkBook.Worksheets.Count; i++)
                {
                    mTempSheet = (Microsoft.Office.Interop.Excel.Worksheet)mWorkBook.Worksheets[i + 1];
                    string sSheetName = mTempSheet.Name;
                    if (sSheetName == sheetName)
                    {
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(to))
            {
                to = from;
            }

            if(mTempSheet == null)
            {
                return string.Empty;
            }

            var range = mTempSheet.get_Range(from, to);
            var cell = range.Cells[1, 1];
            if (cell != null)
            {
                if (cell.Value == null)
                {
                    return string.Empty;
                }
                return cell.Value.ToString();
            }

            return string.Empty;
        }

        /// <summary>
        /// 获取指定位置的值
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string GetRangeValue(RangeData data)
        {
            data.Value = GetRangeValue(data.SheetName, data.From, data.To);
            return data.Value;
        }

        /// <summary>
        /// 设置指定位置的值
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="textValue"></param>
        /// <returns></returns>
        private bool SetRangeValue(string sheetName, string from, string to, string textValue)
        {
            if (mWorkBook == null)
            {
                return false;
            }
            if (mTempSheet == null || mTempSheet.Name != sheetName)
            {
                for (int i = 0; i < mWorkBook.Worksheets.Count; i++)
                {
                    mTempSheet = (Worksheet)mWorkBook.Worksheets[i + 1];
                    string sSheetName = mTempSheet.Name;
                    if (sSheetName == sheetName)
                    {
                        break;
                    }
                }
            }
            if(string.IsNullOrEmpty(to))
            {
                to = from;
            }
            var range = mTempSheet.get_Range(from, to);
            var cell = range.Cells[1, 1];
            if (cell != null)
            {
                cell.Value = textValue;
                return true;
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool SetRangeValue(RangeData data)
        {
            return SetRangeValue(data.SheetName, data.From, data.To, data.Value);
        }

        /// <summary>
        /// 批量写入
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public bool SetRangeValues(IEnumerable<RangeData> datas)
        {
            bool result = true;
            foreach (var data in datas)
            {
                result &=SetRangeValue(data.SheetName, data.From, data.To, data.Value);
            }
            return result;
        }

        /// <summary>
        /// 获取区域集合的信息
        /// </summary>
        /// <param name="rangeDatas"></param>
        /// <returns></returns>
        public bool GetRangesValue(ExcelRangeKeyDatas rangeDatas)
        {
            foreach (var data in rangeDatas)
            {
                GetRangeValue(data);
            }
            return true;
        }

        #endregion .....获取/设定指定位置的值.....

        /// <summary>
        /// 另存为
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            try
            {
                mWorkBook.SaveAs(fileName);
                ExcuteMacro();
                mWorkBook.Close();
            }
            catch (Exception)
            {
                mHasException = true;
            }

        }

        private bool mHasException;

       /// <summary>
       /// 保存完成之后，给控件复制的宏
       /// </summary>
        private const string MacroName = "ThisWorkbook.ControlAssignValues";

        /// <summary>
        /// 执行宏
        /// </summary>
        private void ExcuteMacro()
        {
            //mExcelClass.GetType().InvokeMember(
            //                                             MacroName,
            //                                             System.Reflection.BindingFlags.Default |
            //                                             System.Reflection.BindingFlags.InvokeMethod,
            //                                             null,
            //                                             mExcelClass,
            //                                             null
            //                                          );
            try
            {

                mExcelClass.Run(MacroName, MissingValue, MissingValue,
                    MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue,
                    MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue,
                    MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue,
                    MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue);
            }

            catch(Exception ee)
            {

            } 
        }

        /// <summary>
        /// 读取标准的信息
        /// </summary>
        /// <param name="sheetName"></param>
        public List<T> GetStarandInfoToList<T>(string sheetName) where T : IReadExcelRows, new()
        {
            if (mTempSheet == null || mTempSheet.Name != sheetName)
            {
                if (mTempSheet != null)
                {
                   Marshal.ReleaseComObject(mTempSheet);
                }
                mTempSheet = null;
                for (int i = 0; i < mWorkBook.Worksheets.Count; i++)
                {
                    mTempSheet = (Worksheet)mWorkBook.Worksheets[i + 1];
                    string sSheetName = mTempSheet.Name;
                    if (sSheetName == sheetName)
                    {
                        break;
                    }
                }
            }
            if (mTempSheet == null)
            {
                return null;
            }
            T temp = new T();
            var readColumn = temp.ReadColumnNunbers;
            var noNullColumnIndex = temp.UnNullableColumnIndex;
            int row = 1;
            List<T> tt = new List<T>();
            while (mTempSheet.Cells[row, noNullColumnIndex].Value != null)
            {
                List<string> columnValues = new List<string>();
                for (int column = 1; column <= readColumn; column++)
                {
                    if (mTempSheet.Cells[row, column].Value == null)
                    {
                        columnValues.Add(string.Empty);
                    }
                    else
                    {
                        columnValues.Add(mTempSheet.Cells[row, column].Value.ToString());
                    }
                }
                T test = new T();
                test.AddReadColumnValues(columnValues.ToArray());
                tt.Add(test);
                row++;
            }
            return tt;
        }

        /// <summary>
        /// 释放当前的Excel文件
        /// </summary>
        public void ReleaseExcelFile()
        {


            if (mExcelClass != null)
            {
                try
                {
                    if (!mHasException)
                    {
                        mExcelClass.Workbooks.Close();
                        mExcelClass.Application.Quit();
                        mExcelClass.Quit();
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                catch (Exception)
                {
                    IntPtr t = new IntPtr(mExcelClass.Hwnd); //得到这个句柄，具体作用是得到这块内存入口
                    int k;
                    GetWindowThreadProcessId(t, out k); //得到本进程唯一标志k
                    Process p = Process.GetProcessById(k);
                    p.Kill(); 
                }
            }
        }

        /// <summary>
        /// 析构
        /// </summary>
        public void Dispose()
        {
            ReleaseExcelFile();
        }


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd,out int ID);

        /// <summary>
        /// 是否包含指定的sheet页面
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool ContainsSheet(string sheetName)
        {
            if (mWorkBook == null)
            {
                return false;
            }
            if (mTempSheet == null || mTempSheet.Name != sheetName)
            {
                mTempSheet = null;
                for (int i = 0; i < mWorkBook.Worksheets.Count; i++)
                {
                    mTempSheet = (Worksheet)mWorkBook.Worksheets[i + 1];
                    string sSheetName = mTempSheet.Name;
                    if (sSheetName == sheetName)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
