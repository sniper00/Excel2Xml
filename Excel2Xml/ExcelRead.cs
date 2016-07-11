using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Excel2Xml
{
    class ExcelRead
    {
        OleDbConnection     conn = null;
        OleDbDataAdapter    da = null;
        DataTable           dt = null;

        public List<string> SheetNameList = new List<string>();
        public void OpenFile(string fileName, bool bHdr = true)
        {
            SheetNameList.Clear();
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(fileName);
            if (string.IsNullOrEmpty(fileType))
                return;

            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";" + ";Extended Properties='Excel 8.0;HDR=[{0}];IMEX=1'";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties='Excel 12.0;HDR=[{0}];IMEX=1'";

            string WithHeader = bHdr ? "YES" : "NO";
            connStr = string.Format(connStr, WithHeader);

            DataTable dtSheetName = null;
            try
            {
                if (null != conn)
                {
                    // 关闭连接
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                        da.Dispose();
                        conn.Dispose();
                    }
                }

                // 初始化连接，并打开
                conn = new OleDbConnection(connStr);
                conn.Open();

                // 获取数据源的表定义元数据                        
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                // 初始化适配器
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];

                    if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                    {
                        continue;
                    }
                    SheetNameList.Add(SheetName);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK);
            }
        }

        public DataTable ReadTable(string SheetName)
        {
            if (!SheetNameList.Contains(SheetName))
            {
                return null;
            }

            try
            {
                string sql_F = "Select * FROM [{0}]";
                da.SelectCommand = new OleDbCommand(string.Format(sql_F, SheetName), conn);
                dt = new DataTable();
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK);
            }
            return dt;
        }

        public void CloseExcel()
        {
            if (null != conn)
            {
                // 关闭连接
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
        }
    }
}
