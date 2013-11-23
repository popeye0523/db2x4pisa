using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data;

namespace db2x4pisa
{
    class Program
    {
        static string strConn = string.Empty;
        static void Main(string[] args)
        {
            try
            {
                string currentDirectory = System.Environment.CurrentDirectory;
                Console.WriteLine("当前目录为：" + currentDirectory);

                string configfile = currentDirectory + "\\program.config";
                Console.WriteLine("读取配置文件：" + configfile);

                FileStream fs = new FileStream("program.config", FileMode.Open);

                StreamReader m_streamReader = new StreamReader(fs);

                m_streamReader.BaseStream.Seek(0, SeekOrigin.Begin);
                string ip = "";
                string port = "";
                string datasource = "";
                string username = "";
                string password = "";
                string excel = "";
                string sheet = "";
                string column = "";
                string rows = "";
                string strLine = m_streamReader.ReadLine();
                while (strLine != null && strLine != "")
                {
                    string[] split = strLine.Split('=');
                    switch (split[0].ToLower())
                    {
                        case "ip":
                            ip = split[1];
                            break;
                        case "port":
                            port = split[1];
                            break;
                        case "datasource":
                            datasource = split[1];
                            break;
                        case "username":
                            username = split[1];
                            break;
                        case "password":
                            password = split[1];
                            break;

                        case "excel":
                            excel = split[1];
                            break;
                        case "sheet":
                            sheet = split[1];
                            break;
                        case "column":
                            column = split[1];
                            break;
                        case "rows":
                            rows = split[1];
                            break;
                    }
                    strLine = m_streamReader.ReadLine();
                }

                m_streamReader.Close();
                m_streamReader.Dispose();
                fs.Close();
                fs.Dispose();

                Console.WriteLine("ip=" + ip);
                Console.WriteLine("port=" + port);
                Console.WriteLine("datasource=" + datasource);
                Console.WriteLine("username=" + username);
                Console.WriteLine("password=" + password);

                strConn = "Provider=IBMDADB2;Location=" + ip.Trim() + ":" + port.Trim() + ";Data Source=" + datasource + ";User ID=" + username + ";Password=" + password + ";";
                Console.WriteLine("strConn=" + strConn);

                string excelstr = currentDirectory + "\\" + excel;
                Console.WriteLine("excelstr=" + excelstr);
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(excelstr, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Visible = false;//读Excel不显示出来影响用户体验

                //得到WorkSheet对象
                Console.WriteLine("1");
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(Int32.Parse(sheet));
                Console.WriteLine("1");
                string sql = string.Empty;
                for (int i = 1; i <= Int32.Parse(rows); i++)
                {
                    Microsoft.Office.Interop.Excel.Range r = ws.Cells[i, Int32.Parse(column)];
                    sql = r.Value;
                    if (sql != null && sql.Length != 0)
                    {
                        sql = sql.Trim().Replace(";", "").Replace("；", "").Replace("‘", "'").Replace("’", "'");
                        Console.WriteLine(sql);
                        DataTable dt = new DataTable();
                        try
                        {
                            dt = getDataSet(sql);

                            if (dt != null)
                            {
                                if (dt.Rows.Count == 1)
                                {
                                    Console.WriteLine(dt.Rows[0].ToString());
                                    for (int j = 1; j <= dt.Columns.Count; j++)
                                    {
                                        ws.Cells[i, Int32.Parse(column) + j] = dt.Rows[0][j - 1].ToString();
                                        Console.WriteLine(dt.Rows[0][j - 1].ToString());
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("返回多行值");
                                    ws.Cells[i, Int32.Parse(column) + 1] = "返回多行值";
                                }
                            }
                            else
                            {
                                Console.WriteLine("没有结果");
                                ws.Cells[i, Int32.Parse(column) + 1] = "没有结果";
                            }
                        }
                        catch (Exception ex)
                        {
                            ws.Cells[i, Int32.Parse(column) + 1] = "sql或者db2错误== " + ex.Message;
                        }
                    }
                }
                wb.Save();
                wb.Close(null, null, null);
                app.Workbooks.Close();
                app.Application.Quit();
                app.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                ws = null;
                wb = null;
                app = null;

                GC.Collect(); 

                Console.Write("按任意键结束 . . . ");
                Console.ReadKey(true);
            }
            catch (Exception ex)
            {
                Console.Write("全局程序错误===" + ex.Message);
                Console.ReadKey(true);
            }

            Environment.Exit(0);
        }

        private static DataTable getDataSet(string sql)
        {
            try
                {
            DataTable dt = null;
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {

                OleDbCommand cmd = new OleDbCommand(sql, conn);
                
                    conn.Open();
                    OleDbDataAdapter adp = new OleDbDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    adp.Fill(ds);
                    dt = ds.Tables[0];
                
            return dt;
        }
                }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
        }


    }
}
