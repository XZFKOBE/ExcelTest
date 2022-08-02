using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using Dapu.DataAccess;
using Dapu.Logging;
using ExcelHelper;
using System.Threading;
namespace ExcelTest
{
    public class SqlServercon
    {
        public int SqlServerTest(string SqlName)
        {

            //string sqlConnectionString = @"Server=192.168.4.9;Uid=dp;Pwd=dp123456;DataBase=KaiTest";
            string localDBConnectionString = @"Server=192.168.4.9;Uid=dp;Pwd=dp123456;DataBase="+ SqlName;
            int sql_ServeState;
            try
            {
                SqlConnection conn = new SqlConnection(localDBConnectionString);
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    sql_ServeState = 1;
                    conn.Close();
                    return sql_ServeState;
                }
                else
                {
                    sql_ServeState = 0;
                    conn.Close();
                    return sql_ServeState;
                }

            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
            return -1;
        }

        /// <summary>
        /// 读取Excel文件特定名字sheet的内容到DataTable
        /// </summary>
        /// <param name="strFileName">excel文件路径</param>
        /// <param name="sheet">需要导出的sheet</param>
        /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
        /// <param name="dir">excel列名和DataTable列名的对应字典</param>
        /// <returns></returns>
        public DataTable ShowTable(string fileName, Dictionary<string, string> dir)
        {
            ExcelDataHelper objExcel = new ExcelDataHelper();

            DataTable Msg = objExcel.ImportExceltoDt(fileName, dir, -1, 0);

            return Msg;
        }

        /// <summary>
        /// 在数据库里创建表
        /// </summary>
        /// <param name="sheetName">要创建的表名</param>
        /// <param name="SqlName">数据库名</param>
        public void CreateTable(string sheetName, string SqlName,DataTable trunSheet)
        {
            int cellCount = trunSheet.Columns.Count;
            string str1 = "CREATE table [" + sheetName + "]";
            string str2 = "(";
            
            for (int i = 0; i < cellCount-1; i++)
            {
                str2 = str2 + "[" + trunSheet.Columns[i].ColumnName + "]" + "VARCHAR(255)" + ",";
            }
            str2 += "[" + trunSheet.Columns[cellCount-1].ColumnName + "]" + "VARCHAR(255));";
            string CreateTable = str1 + str2;
            // Console.WriteLine(CreateTable);

            try
            {
                string localDBConnectionString = @"Server=192.168.4.9;Uid=dp;Pwd=dp123456;DataBase=" + SqlName;
                SqlConnection conn = new SqlConnection(localDBConnectionString);
                conn.Open();
                //Console.WriteLine(CreateTable);
                SqlCommand mycmd = new SqlCommand(CreateTable, conn);
                mycmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
            /*81行，0-94.95列. cellCount = 94*/
            
        }
        /// <summary>
        /// 将dataTable的内容读到数据库表中
        /// </summary>
        /// <param name="sheet">要录入的表</param>
        /// <param name="sheetName">数据库中的表名</param>
        /// <param name="sqlName">数据库名</param>
        public void WriteSql_5699c(DataTable trunSheet, string sheetName, string SqlName)
        {
            int cellCount = trunSheet.Columns.Count;
            string localDBConnectionString = @"Server=192.168.4.9;Uid=dp;Pwd=dp123456;DataBase=" + SqlName;
            string adddata = "";
            string insertStr = "insert into  [" + sheetName + "]";
            string title = "(";
            for (int i = 0; i < cellCount-1; i++) 
            {
                title = title + "[" + trunSheet.Columns[i].ColumnName + "]" + ",";
            }
            title += "[" + trunSheet.Columns[cellCount - 1].ColumnName + "]" + ")" + "values";

            //adddata = insertStr + title;
            //Console.WriteLine(adddata);
            string date = "(";
            for (int i = 0; i < 8; i++)
            {
                date += "'" + trunSheet.Rows[0][i].ToString() + "'" + ",";
            }
            string fixedDate = date;
           // Console.WriteLine(fixedDate);
            string sumData = "";
            int trunSheetRow = trunSheet.Rows.Count;
            int line = 0;
            int temp = 0;
            SqlConnection write = new SqlConnection(localDBConnectionString);
            write.Open();
            for (int i = 0; i < trunSheetRow ; i++)//行遍历
            {
                
                for (int j = 8; j < cellCount - 1; j++)//列遍历
                {
                    string tableNum = trunSheet.Rows[i].ItemArray[j].ToString();
                    if (tableNum == "")
                    {
                        date += "null" + ",";
                    }
                    else
                    {
                        date += "'" + tableNum + "'" + ",";
                    }

                }
                if (trunSheet.Rows[i].ItemArray[cellCount - 1].ToString() == "")
                {
                    date += "null" + "),";
                }
                else
                {
                    date += "'" + trunSheet.Rows[i].ItemArray[cellCount - 1].ToString() + "'" + "),";
                }
                sumData += date;
                temp++;

                //Console.WriteLine(sumData);

                //adddata = insertStr + title + date;
                //Console.WriteLine(adddata);
                //sumData += adddata;
                //SqlCommand writeSql = new SqlCommand(adddata, write);
                //line =  writeSql.ExecuteNonQuery();
                if (temp == 1000)
                {
                    sumData = sumData.Substring(0, sumData.Length - 1);
                    sumData += ";";
                    
                    adddata = insertStr + title + sumData;


                    SqlCommand writeSql1 = new SqlCommand(adddata, write);
                    writeSql1.CommandTimeout = 200;
                   Console.WriteLine(adddata);
                    writeSql1.ExecuteNonQuery();
                    //Console.WriteLine(sumData);
                    //Console.WriteLine(sumData);
                    //write.Close();
                    temp = 0;
                    sumData = "";
                }
                adddata = "";
                date = fixedDate;
                
                
            }
           // write.Close();
            temp = 0;
            //Console.WriteLine (sumData.Length.ToString());
            sumData =  sumData.Substring(0, sumData.Length - 1);
            sumData += ";";
            adddata = insertStr + title + sumData;

            Console.WriteLine(adddata);
            
            //write2.Open();
            SqlCommand writeSql2 = new SqlCommand(adddata, write);
            writeSql2.CommandTimeout = 200;
            writeSql2.ExecuteNonQuery();
            //Console.WriteLine(sumData);
            //Console.WriteLine(sumData);
            
            write.Close();
        }

        public DataTable Trunsheet(DataTable sheet)
        {
            DataTable trunSheet = new DataTable();
            //dataGridView1.DataSource = sheet;
            for (int i = 0; i < 8; i++)//添加前八个字段(列头)
            {

                trunSheet.Columns.Add(sheet.Rows[i][0].ToString(), typeof(String));//列
                trunSheet.Rows.Add();//添加行
                trunSheet.Rows[0][i] = sheet.Rows[i][1].ToString();//同一行不同列

            }
            trunSheet.Columns.Add("Test_time", typeof(String));//列
            int timeNum = 0;
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                bool timeStatu = sheet.Rows[i][0].ToString().Contains("2022"); //随年份改动即可
                if (timeStatu == true)
                {

                    timeNum++;

                }
            }
            for (int i = 0; i < timeNum - 8; i++)//添加行,test_time的时间行为参照行

            {
                trunSheet.Rows.Add();//添加行

            }
            for (int i = 11, j = 0; i < sheet.Rows.Count && j < timeNum; i++, j++)//写test 时间
            {

                trunSheet.Rows[j]["Test_time"] = sheet.Rows[i][0].ToString();

            }

            int cellCount;
            cellCount = ExcelDataHelper.LastCellCount;//添加Test_Count 到最后一列的所有列
            for (int i = 1; i < cellCount - 6; i++)
            {
                //Console.WriteLine(sheet.Rows[8][i].ToString());
                trunSheet.Columns.Add(sheet.Rows[8][i].ToString(), typeof(String));//列
            }


            for (int k = 1; k < cellCount - 6; k++)
            {
                string testName = sheet.Rows[8][k].ToString();
                //Console.WriteLine(testName);
                for (int i = 11, j = 0; i < sheet.Rows.Count && j < timeNum; i++, j++)
                {

                    trunSheet.Rows[j][testName] = sheet.Rows[i][k].ToString();

                }
            }
            return trunSheet;

        }

        public DataTable WriteIntoTestresult(DataTable TrunSheet, string resultSheet, string SqlServer)//筛选字段写进t_rtc_ft_testresult数据库
        {
            DataTable testresult = new DataTable();
            testresult.Columns.Add("RtcID", typeof(String));//列
            testresult.Columns.Add("TestTime", typeof(String));//列
            testresult.Columns.Add("IDD7", typeof(String));//列
            testresult.Columns.Add("IDD8", typeof(String));//列
            testresult.Columns.Add("IDD9", typeof(String));//列
            testresult.Columns.Add("IDD10", typeof(String));//列
            testresult.Columns.Add("IDD11", typeof(String));//列
            testresult.Columns.Add("IDD12", typeof(String));//列
            testresult.Columns.Add("IDD13", typeof(String));//列
            testresult.Columns.Add("IDD14", typeof(String));//列
            testresult.Columns.Add("IDD15", typeof(String));//列
            testresult.Columns.Add("IDD16", typeof(String));//列
            testresult.Columns.Add("Frequency", typeof(String));//列
            testresult.Columns.Add("TrackingCardId", typeof(String));//列
            testresult.Columns.Add("ProductType", typeof(String));//列
            testresult.Columns.Add("EquipmentName", typeof(String));//列


            for (int i = 0; i < TrunSheet.Rows.Count; i++)//添加行,trunSheet的表行为参照行

            {
                testresult.Rows.Add();//添加行

            }
            for (int i = 0; i < TrunSheet.Rows.Count; i++)// 填充"RtcID" "TestTime" "Frequency"  "TrackingCardId"行数据
            {
                testresult.Rows[i]["RtcID"] = TrunSheet.Rows[i]["RTCID"].ToString();
                testresult.Rows[i]["TestTime"] = TrunSheet.Rows[i]["Test_time"].ToString();
                testresult.Rows[i]["Frequency"] = TrunSheet.Rows[i]["FREQ_Value"].ToString();
                testresult.Rows[i]["TrackingCardId"] = TrunSheet.Rows[0]["LOT_ID:"].ToString();
                testresult.Rows[i]["ProductType"] = TrunSheet.Rows[0]["DEVICE:"].ToString();
                testresult.Rows[i]["EquipmentName"] = "利扬";


            }
            for (int i = 7; i < 17; i++)//填充IDD7-IDD16 数据
            {
                for (int j = 0; j < TrunSheet.Rows.Count; j++)
                {
                    testresult.Rows[j]["IDD" + i.ToString()] = TrunSheet.Rows[j]["IDD" + i.ToString()].ToString();
                }
            }

            for (int i = 0; i < TrunSheet.Rows.Count; i++)//改写时间格式，与数据库中datetime 格式对应
            {
                string time = testresult.Rows[i]["TestTime"].ToString();
                string[] timeArray = time.Split('-');
                testresult.Rows[i]["TestTime"] = timeArray[0] + "-" + timeArray[1] + "-" + timeArray[2] + " " + timeArray[3];
                time = "";
                timeArray = null;
            }
            string localDBConnectionString = @"Server=192.168.4.9;Uid=dp;Pwd=dp123456;DataBase=" + SqlServer;
            string adddata = "";
            string insertStr = "insert into  " + resultSheet;
            string title = "(";
            int cellCount = testresult.Columns.Count;
            int temp = 0;
            int line = 0;
            for (int i = 0; i < cellCount - 1; i++)
            {
                title = title + testresult.Columns[i].ColumnName + ",";
            }
            title += testresult.Columns[cellCount - 1].ColumnName + ")" + "values";
            //Console.WriteLine(insertStr + title);
            string date = "(";
            int trunSheetRow = testresult.Rows.Count;
            string sumData = "";
            SqlConnection write = new SqlConnection(localDBConnectionString);
            write.Open();
            for (int i = 0; i < trunSheetRow; i++)//行遍历
            {

                for (int j = 0; j < cellCount - 1; j++)//列遍历
                {
                    string tableNum = testresult.Rows[i].ItemArray[j].ToString();

                    date += "'" + tableNum + "'" + ",";


                }

                date += "'" + testresult.Rows[i].ItemArray[cellCount - 1].ToString() + "'" + "),";
                sumData += date;
                temp++;

                if (temp == 1000)
                {
                    sumData = sumData.Substring(0, sumData.Length - 1);
                    sumData += ";";

                    adddata = insertStr + title + sumData;


                    SqlCommand writeSql1 = new SqlCommand(adddata, write);
                    writeSql1.CommandTimeout = 200;
                    Console.WriteLine(adddata);
                    writeSql1.ExecuteNonQuery();
                    //Console.WriteLine(sumData);
                    //Console.WriteLine(sumData);
                    //write.Close();
                    temp = 0;
                    sumData = "";
                }                
                adddata = "";
                date = "(";
            }

            temp = 0;
            //Console.WriteLine (sumData.Length.ToString());
            sumData = sumData.Substring(0, sumData.Length - 1);
            sumData += ";";
            adddata = insertStr + title + sumData;

            Console.WriteLine(adddata);

            //write2.Open();
            SqlCommand writeSql2 = new SqlCommand(adddata, write);
            writeSql2.CommandTimeout = 200;
            writeSql2.ExecuteNonQuery();
            //Console.WriteLine(sumData);
            //Console.WriteLine(sumData);

            write.Close();
            return testresult;
        }

        public DataTable WriteSql_5710A(DataTable sheet, string sheetName, string SqlName)
        {

            DataTable trunSheet = new DataTable();
            for (int i = 0; i < 8; i++)//添加前八个字段(列头)
            {

                trunSheet.Columns.Add(sheet.Rows[i][0].ToString(), typeof(String));//列
                trunSheet.Rows.Add();//添加行
                trunSheet.Rows[0][i] = sheet.Rows[i][1].ToString();//同一行不同列

            }
            trunSheet.Columns.Add("Test_time", typeof(String));//列

            int timeNum = 0;
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                bool timeStatu = sheet.Rows[i][0].ToString().Contains("2022"); //随年份改动即可
                if (timeStatu == true)
                {

                    timeNum++;

                }
            }
            for (int i = 0; i < timeNum - 8; i++)//添加行,test_time的时间行为参照行

            {
                trunSheet.Rows.Add();//添加行

            }

            for (int i = 11, j = 0; i < sheet.Rows.Count && j < timeNum; i++, j++)//写test 时间
            {

                trunSheet.Rows[j]["Test_time"] = sheet.Rows[i][0].ToString();

            }

            int cellCount;
            cellCount = ExcelDataHelper.LastCellCount;//添加Test_Count 到表的最后一列的所有列
            for (int i = 1; i <= cellCount; i++)
            {
                //Console.WriteLine(sheet.Rows[8][i].ToString());
                trunSheet.Columns.Add(sheet.Rows[8][i].ToString(), typeof(String));//列
            }


            for (int k = 1; k <= cellCount; k++)//填充数据
            {
                string testName = sheet.Rows[8][k].ToString();
                //Console.WriteLine(testName);
                for (int i = 11, j = 0; i < sheet.Rows.Count && j < timeNum; i++, j++)
                {

                    trunSheet.Rows[j][testName] = sheet.Rows[i][k].ToString();
                    trunSheet.Rows[j]["DEVICE:"] = sheet.Rows[2][1].ToString();//原表DEVICE:
                    trunSheet.Rows[j]["IntDevice:"] = sheet.Rows[2][1].ToString();//
                    trunSheet.Rows[j]["Customer:"] = sheet.Rows[0][1].ToString();//
                    trunSheet.Rows[j]["PO_NO:"] = sheet.Rows[3][1].ToString();//
                    trunSheet.Rows[j]["LOT_ID:"] = sheet.Rows[4][1].ToString();//
                    trunSheet.Rows[j]["Program:"] = sheet.Rows[5][1].ToString();//
                    trunSheet.Rows[j]["FT/RT:"] = sheet.Rows[7][1].ToString();//

                }

            }
            string localDBConnectionString = @"Server=192.168.4.9;Uid=dp;Pwd=dp123456;DataBase=" + SqlName;
            SqlConnection cellTest = new SqlConnection(localDBConnectionString);

            cellTest.Open();
            for (int i = 0; i < trunSheet.Columns.Count; i++)//判断数据库的表中是否存在某个字段名，若无则添加
            {
                string columnName = trunSheet.Columns[i].ColumnName;
                string findColumnName = "select * from syscolumns where id=object_id('" + sheetName + "') and name='" + columnName + "'";
                string addColumName = "ALTER TABLE [" + sheetName + "] ADD [" + columnName + "] VARCHAR(255)";
                SqlDataAdapter writeSql1 = new SqlDataAdapter(findColumnName, cellTest);
                DataSet ds = new DataSet();
                writeSql1.Fill(ds);
                //Console.WriteLine(ds.Tables[0].Rows.Count.ToString());
                //Console.WriteLine(columnName);
                if (ds.Tables[0].Rows.Count <= 0)
                {
                    SqlCommand addColum = new SqlCommand(addColumName, cellTest);
                    addColum.ExecuteNonQuery();
                }
            }
            cellTest.Close();
            string adddata = "";
            string insertStr = "insert into  [" + sheetName + "]";
            string title = "(";
            int truncellCount = trunSheet.Columns.Count;
            int temp = 0;
            int line = 0;
            for (int i = 0; i < truncellCount - 1; i++)
            {
                title = title + "[" + trunSheet.Columns[i].ColumnName + "]" + ",";
            }
            title += "[" + trunSheet.Columns[truncellCount - 1].ColumnName + "]" + ")" + "values";
            //Console.WriteLine(insertStr + title);
            string date = "(";
            //Console.WriteLine(insertStr + title + date);
            int trunSheetRow = trunSheet.Rows.Count;
            string sumData = "";
            SqlConnection write = new SqlConnection(localDBConnectionString);
            write.Open();
            for (int i = 0; i < trunSheetRow; i++)//行遍历
            {

                for (int j = 0; j < truncellCount - 1; j++)//列遍历
                {
                    string tableNum = trunSheet.Rows[i].ItemArray[j].ToString();
                    if (tableNum == "")
                    {
                        date += "null" + ",";
                    }
                    else
                    {
                        date += "'" + tableNum + "'" + ",";
                    }

                }
                if (trunSheet.Rows[i].ItemArray[truncellCount - 1].ToString() == "")
                {
                    date += "null" + "),";
                }
                else
                {
                    date += "'" + trunSheet.Rows[i].ItemArray[truncellCount - 1].ToString() + "'" + "),";
                }
                sumData += date;
                //Console.WriteLine(insertStr +title+ sumData);
                temp++;

                if (temp == 1000)
                {
                    sumData = sumData.Substring(0, sumData.Length - 1);
                    sumData += ";";

                    adddata = insertStr + title + sumData;


                    SqlCommand writeSql1 = new SqlCommand(adddata, write);
                    writeSql1.CommandTimeout = 200;
                    Console.WriteLine(adddata);
                    writeSql1.ExecuteNonQuery();
                    //Console.WriteLine(sumData);
                    //Console.WriteLine(sumData);
                    //write.Close();
                    temp = 0;
                    sumData = "";
                }
                adddata = "";
                date = "(";
            }

            temp = 0;
            //Console.WriteLine (sumData.Length.ToString());
            sumData = sumData.Substring(0, sumData.Length - 1);
            sumData += ";";
            adddata = insertStr + title + sumData;

            Console.WriteLine(adddata);

            //write2.Open();
            SqlCommand writeSql2 = new SqlCommand(adddata, write);
            writeSql2.CommandTimeout = 200;
            writeSql2.ExecuteNonQuery();
            //Console.WriteLine(sumData);
            //Console.WriteLine(sumData);

            write.Close();
            return trunSheet;
        }

    }
 }

