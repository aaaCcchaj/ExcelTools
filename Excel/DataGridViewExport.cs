using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FS.Tool.Excel
{
    public class DataGridViewExport
    {
         public static void DataToExcel(string defaultName,DataGridView dataView)
         {
              SaveFileDialog kk = new SaveFileDialog();
              kk.Title = "保存EXECL文件"; 
              kk.Filter = "EXECL文件(*.xls) |*.xls |所有文件(*.*) |*.*"; 
              kk.FilterIndex = 1;
             kk.FileName = defaultName;
             if (kk.ShowDialog() == DialogResult.OK) 
              { 
                  string FileName = kk.FileName ;
                 if (File.Exists(FileName))
                     File.Delete(FileName);
                 FileStream objFileStream; 
                 StreamWriter objStreamWriter; 
                 string strLine = ""; 
                 objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write); 
                 objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);
                 for (int i = 0; i  < dataView.Columns.Count; i++) 
                 { 
                     if (dataView.Columns[i].Visible == true) 
                     { 
                        strLine = strLine + dataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9); 
                     } 
                 } 
                 objStreamWriter.WriteLine(strLine); 
                 strLine = ""; 
 
                 for (int i = 0; i  < dataView.Rows.Count; i++) 
                 { 
                     if (dataView.Columns[0].Visible == true) 
                     { 
                         if (dataView.Rows[i].Cells[0].Value == null) 
                             strLine = strLine + " " + Convert.ToChar(9); 
                         else 
                             strLine = strLine + dataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9); 
                    } 
                     for (int j = 1; j  < dataView.Columns.Count; j++) 
                   { 
                     if (dataView.Columns[j].Visible == true) 
                       { 
                            if (dataView.Rows[i].Cells[j].Value == null) 
                                strLine = strLine + " " + Convert.ToChar(9); 
                            else 
                            { 
                                string rowstr = ""; 
                                rowstr = dataView.Rows[i].Cells[j].Value.ToString(); 
                                if (rowstr.IndexOf("\r\n") >  0) 
                                    rowstr = rowstr.Replace("\r\n", " "); 
                                if (rowstr.IndexOf("\t") >  0) 
                                    rowstr = rowstr.Replace("\t", " "); 
                                strLine = strLine + rowstr + Convert.ToChar(9); 
                            } 
                        } 
                    } 
                    objStreamWriter.WriteLine(strLine); 
                    strLine = ""; 
                } 
                objStreamWriter.Close(); 
                objFileStream.Close();
                MessageBox.Show("保存EXCEL成功","提示",MessageBoxButtons.OK,MessageBoxIcon.Information); 
            }
        }
 }
}
