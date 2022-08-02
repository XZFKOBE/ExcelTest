using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Dapu.Logging;
using System.IO;
using ExcelHelper;
using MySql.Data.MySqlClient;
using System.Threading;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        string FileName = "";
        string SqlServer = "DPQMSRTC";//数据库名
        string sheetName = "t_rtc_ft_testresult_ly";//数据库中的表名
        //string resultSheet = "t_rtc_ft_testresult_ly_test";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//选择文件
        {
            OpenFileDialog ofd = new OpenFileDialog();//新建打开文件对话框
            ofd.InitialDirectory = Directory.GetCurrentDirectory();//设置初始文件目录
            //ofd.Filter = "Excel文档(*.xls)|*.xls|Excel2007文档(*.xlsx)|*.xlsx";//设置打开文件类型
            ofd.Filter = "Excel2007文档(*.xlsx)|*.xlsx";//设置打开文件类型
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                FileName = ofd.FileName;//FileName就是要打开的文件路径
                textBox1.Text = FileName;
                textBox3.Text = Path.GetFileNameWithoutExtension(FileName);
                //Console.WriteLine(textBox1.Text);
            }
        }



        private void button2_Click(object sender, EventArgs e)//导入
        {

            Thread tips = new Thread(Tips);
            tips.IsBackground = true;
            tips.Start();
            button2.Visible = false;
            Dictionary<string, string> dir = new Dictionary<string, string>();
            // dir.Add("Customer:", "Customer:");

            SqlServercon objSheet = new SqlServercon();
            DataTable sheet = objSheet.ShowTable(FileName, dir);//excel导入dataTable
            dataGridView1.DataSource = sheet;
            DataTable trunSheet = new DataTable();

            trunSheet = objSheet.Trunsheet(sheet);//将Excel的表格转换为需要的表结构

            //C:\Users\Asus\Desktop\5710A_XB\5710A_XB-Z22022300030-FT20220406164-FT1-FT1-20220417141115--1.xlsx
            //Console.WriteLine(trunSheet.Rows[0][0].ToString());
            //Console.WriteLine(trunSheet.Rows.Count.ToString());
            //Console.WriteLine(trunSheet.Columns.Count.ToString());
            //Console.WriteLine(trunSheet.Columns[0].ColumnName);
            //在数据库中创表
            try
            {
                objSheet.CreateTable(sheetName, SqlServer, trunSheet);
                objSheet.WriteSql_5699c(trunSheet, sheetName, SqlServer);//写表                             
                //objSheet.WriteIntoTestresult(trunSheet, resultSheet, SqlServer);//写进t_rtc_ft_testresult表中
                //dataGridView1.DataSource = objSheet.WriteIntoTestresult(trunSheet, resultSheet, SqlServer);//写进t_rtc_ft_testresult表中
                button2.Visible = true;
                dataGridView1.DataSource = trunSheet;
                tips.Abort();
                MessageBox.Show("成功将数据写入数据库！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                tips.Abort();
                MessageBox.Show("写入失败：", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

            }
        }
        public void Tips()
        {
            while (true)
            {
                MessageBox.Show("正在导入数据库，请不要进行其他操作,请等待", "导入中", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }

        }


        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)//测试连接
        {

            SqlServercon mysqlObj = new SqlServercon();
            int SqlStatus = mysqlObj.SqlServerTest(SqlServer);
            switch (SqlStatus)
            {
                case 1:
                    MessageBox.Show("数据库连接成功");
                    break;
                default:
                    MessageBox.Show("连接失败，请重试");
                    break;
            }
           

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void mySqlconBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Thread tips = new Thread(Tips);
            tips.IsBackground = true;
            tips.Start();
            button4.Visible = false;
            Dictionary<string, string> dir = new Dictionary<string, string>();
            SqlServercon objSheet = new SqlServercon();
            DataTable sheet = objSheet.ShowTable(FileName, dir);//excel导入dataTable
            try
            {
                DataTable trunSheet =  objSheet.WriteSql_5710A(sheet, sheetName, SqlServer);
                button4.Visible = true;
                dataGridView1.DataSource = trunSheet;
                tips.Abort();
                MessageBox.Show("成功将数据写入数据库！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
             catch (Exception ex)
            {
                button4.Visible = true;
                Logger.Error(ex);
                tips.Abort();
                MessageBox.Show("写入失败：", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

            }
            //dataGridView1.DataSource = sheet;


            
        }
    }
 }

