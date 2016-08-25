using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FillExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // .CSV
            #region  当文件是 csv 格式的文件时 使用的代码
            // 源文件路径  自行修改
            //string sourcePath = @"C:\Users\Petrel\address.csv";


            //// 如果要用命令行参数 动态传进源文件路径 请注释掉上一行  取消下一行的注释掉
            //// string sourcePath = args[1];  

            //// 目标文件路径 自行修改
            //string targetPath = @"C:\Users\Petrel\tar-address.csv";

            //// 命令行动态传入目标文件路径 请注释掉上一行  取消下一行的注释
            //// string targetPath = args[2];

            //FileStream sourceFs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read);

            //// 打开源文件
            //StreamReader sr = new StreamReader(sourceFs, Encoding.GetEncoding("utf-8"));
            //StringBuilder sb = new StringBuilder();

            //string line = string.Empty;
            //line = sr.ReadLine();


            //// 遍历每行  读取数据
            //#region 遍历每行

            //// 三个临时变量
            //string tempPro = string.Empty; // 临时省份
            //string tempCity = string.Empty; // 临时城市
            //string tempCounty = string.Empty; // 临时区县

            //while (line != null)
            //{
            //    // 读取每行 并分割成数组 
            //    string[] dataline = line.Split(',');

            //    if (!string.IsNullOrEmpty(dataline[1]))
            //    {

            //        // 如果判断出省份 不为空 则重置临时变量 tempCity tempCounty
            //        // 不为空 则说明 到了新的省份了
            //        // 然后单独拼接这一行

            //        tempPro = dataline[1];
            //        tempCity = string.Empty;
            //        tempCounty = string.Empty;

            //        // 拼接新一行
            //        Console.WriteLine("Catching the Column Proivce is not null, and now changing the provice");
            //        sb.Append(AppendData(dataline));
            //    }
            //    else
            //    {
            //        if (!string.IsNullOrEmpty(dataline[2]))
            //        {
            //            // 判断 城市 City列是否不为空
            //            // 不为空则说明 换 城市了  要更新  tempCity 这个临时变量
            //            // 然后重置  区县 tempCounty 这个临时变量
            //            // 这里仍然要单独拼接

            //            tempCity = dataline[2];
            //            tempCounty = string.Empty;
            //            dataline[1] = tempPro;
            //            // 拼接到新的一行

            //            Console.WriteLine("Catching the Column City is not null, and now changing the city");
            //            sb.Append(AppendData(dataline));
            //        }
            //        else
            //        {
            //            if (!string.IsNullOrEmpty(dataline[3]))
            //            {
            //                // 判断 区县 County 列 是否不为空
            //                // 不为空 说明到了 新的乡镇  要更新 tempCounty 这个临时变量
            //                // 然后修改 dataline 中的 省份 和 城市 的值

            //                tempCounty = dataline[3];
            //                dataline[1] = tempPro;
            //                dataline[2] = tempCity;

            //                Console.WriteLine("Catching the Column County is not null, and now changing the county");
            //                sb.Append(AppendData(dataline));
            //            }
            //            else
            //            {
            //                dataline[1] = tempPro;
            //                dataline[2] = tempCity;
            //                dataline[3] = tempCounty;

            //                sb.Append(AppendData(dataline));
            //            }
            //        }
            //    }

            //    line = sr.ReadLine();
            //}
            //#endregion

            //// 遍历完了之后 清理缓存 关闭读取流 和 文件流 

            //Console.WriteLine("Iteration is finishing");
            //Console.WriteLine("Now ending the streamreader");

            //sourceFs.Flush();
            //sourceFs.Close();
            //sr.Close();

            //Console.WriteLine("Starting a new Stream");

            //// 遍历完了之后 就打开新文件
            //// 重新写入
            //FileStream targetFs = new FileStream(targetPath, FileMode.OpenOrCreate, FileAccess.Write);

            //Console.WriteLine("Staring a Stream Writer");
            //StreamWriter sw = new StreamWriter(targetFs,Encoding.UTF8);

            //Console.WriteLine("Writing....");
            //// 写入目标文件
            //sw.WriteLine(sb.ToString());

            //Console.WriteLine("Everything is OK");

            //sw.Close();
            //targetFs.Close();


            //Console.WriteLine("Closing the File Stream");


            //Console.ReadKey();

            #endregion

            // .excel
            #region 当文件为 xls 或者是 xlsx 格式的文件时  使用的代码

            // 源文件路径
            string sourcePath = @"C:\Users\Petrel\address.xlsx";

            // 目标文件路径
            string targetPath = @"C:\Users\Petrel\new-address.csv";

            ExcelType type = ExcelType.XLS;
            if (sourcePath.Split('.')[1] == "xlsx")
            {
                type = ExcelType.XLSX;
            }

            // 从 Excel 获取数据
            #region 读取数据到 DataTable 
            DataTable dt = ReadExcelToTable(sourcePath, type);

            if (dt == null)
            {
                Console.WriteLine("Can't not open the excel file");
                Console.Read();

                return;
            }
            // 三个临时变量
            string tempPro = string.Empty; // 临时省份
            string tempCity = string.Empty; // 临时城市
            string tempCounty = string.Empty; // 临时区县

            StringBuilder sb = new StringBuilder();
            
            foreach (DataRow item in dt.Rows)
            {
                string[] dataline = new string[item.ItemArray.Length];
                for (int i = 0; i < item.ItemArray.Length; i++)
                {
                    dataline[i] = item.ItemArray[i].ToString();
                }

                if (!string.IsNullOrEmpty(dataline[1]))
                {
                    // 如果判断出省份 不为空 则重置临时变量 tempCity tempCounty
                    // 不为空 则说明 到了新的省份了
                    // 然后单独拼接这一行

                    tempPro = dataline[1];
                    tempCity = string.Empty;
                    tempCounty = string.Empty;

                    // 拼接新一行
                    Console.WriteLine("Catching the Column Proivce is not null, and now changing the provice");
                    sb.Append(AppendData(dataline));
                }
                else
                {
                    if (!string.IsNullOrEmpty(dataline[2]))
                    {
                        // 判断 城市 City列是否不为空
                        // 不为空则说明 换 城市了  要更新  tempCity 这个临时变量
                        // 然后重置  区县 tempCounty 这个临时变量
                        // 这里仍然要单独拼接

                        tempCity = dataline[2];
                        tempCounty = string.Empty;
                        dataline[1] = tempPro;
                        // 拼接到新的一行

                        Console.WriteLine("Catching the Column City is not null, and now changing the city");
                        sb.Append(AppendData(dataline));
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(dataline[3]))
                        {
                            // 判断 区县 County 列 是否不为空
                            // 不为空 说明到了 新的乡镇  要更新 tempCounty 这个临时变量
                            // 然后修改 dataline 中的 省份 和 城市 的值

                            tempCounty = dataline[3];
                            dataline[1] = tempPro;
                            dataline[2] = tempCity;

                            Console.WriteLine("Catching the Column County is not null, and now changing the county");
                            sb.Append(AppendData(dataline));
                        }
                        else
                        {
                            dataline[1] = tempPro;
                            dataline[2] = tempCity;
                            dataline[3] = tempCounty;

                            sb.Append(AppendData(dataline));
                        }
                    }
                }
            }
            #endregion

            // 写数据到 CSV
            #region 写数据到 CSV
            
            // 重新写入
            FileStream targetFs = new FileStream(targetPath, FileMode.OpenOrCreate, FileAccess.Write);

            Console.WriteLine("Staring a Stream Writer");
            StreamWriter sw = new StreamWriter(targetFs, Encoding.UTF8);

            Console.WriteLine("Writing....");
            // 写入目标文件
            sw.WriteLine(sb.ToString());

            Console.WriteLine("Everything is OK");

            sw.Close();
            targetFs.Close();

            Console.WriteLine("Closing the File Stream");

            Console.ReadKey();

            #endregion


            #endregion
        }

        /// <summary>
        /// 拼接字符串
        /// </summary>
        /// <param name="data"></param>
        /// <param name="sb"></param>
        /// <returns></returns>
        static StringBuilder AppendData(string[] data)
        {
            StringBuilder _sb = new StringBuilder();
            _sb.Append("\r\n");
            foreach (var item in data)
            {
                _sb.Append(item + ",");
            }

            return _sb;
        }

        /// <summary>
        /// 读取 Excel 表格 到 DataTable
        /// </summary>
        /// <param name="path"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        static DataTable ReadExcelToTable(string path, ExcelType type)
        {
            string connStr = string.Empty;
            //
            // Excel 有两个不同内核版本
            // office 2007 以上 包括 2007 文件名后缀 .xlsx
            // office 2007 以下 文件名后缀为 .xls
            // 两个版本的驱动文件不同 
            // 这里 要判断 是哪个版本  选择 不同的驱动
            //
            if (type == ExcelType.XLS)
            {
                connStr = "Provider=Microsoft.JET.OLEDB.4.0;DataSource=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            }
            else if (type == ExcelType.XLSX)
            {
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            }

            try
            {
                DataSet set = new DataSet();

                // 使用 OleDb 打开文件 把excel文件当做数据库来用
                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    DataTable sheetDT = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });

                    // 获取excel表里面第一个 sheet 的名字
                    string firstSheet = sheetDT.Rows[0][2].ToString();

                    string sql = string.Format("select * from [{0}]", firstSheet);

                    OleDbCommand comm = new OleDbCommand(sql);
                    comm.Connection = conn;
                    OleDbDataReader reader = comm.ExecuteReader();
                    while (reader.Read())
                    {
                        Console.WriteLine(reader.GetString(0));
                        Console.WriteLine(reader.GetString(1));
                        Console.WriteLine(reader.GetString(2));
                        Console.WriteLine(reader.GetString(3));
                    }

                    OleDbDataAdapter da = new OleDbDataAdapter(sql, connStr);

                    da.Fill(set);
                }
                // 返回 获取的数据 DataTable
                return set.Tables[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);

                Console.WriteLine("Something wrong!");

                return null;
            }

        }
    }

    /// <summary>
    /// excel 表格的格式类型
    /// </summary>
    enum ExcelType
    {
        XLS,
        XLSX
    }
}
