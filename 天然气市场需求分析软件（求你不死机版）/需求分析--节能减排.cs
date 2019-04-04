using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows6 : Form
    {
        public Windows6()
        {
            InitializeComponent();
        }
        private void Add1()
        {
            try
            {
                string str1= lblInput1.Text;
                string str2 = txtInput1.Text;
                Common.ParameterErrorDetectionNaturalGasConsume(str1, str2);
                double Input1 = Convert.ToDouble(txtInput1.Text);  //输入量
                txtOutput1.Text  = (Input1 * 20.2).ToString("0.0");
                txtOutput2.Text = (Input1 * 7.6).ToString("0.0");
                txtOutput3.Text = (Input1 *37.3).ToString("0.0");
                txtOutput4.Text = (Input1 * 0.2561).ToString("0.0");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void txtInput1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
                Add1();
         
        }
        private void Clear()
        {
            txtInput1.Text = "";
            txtOutput1.Text = "";
            txtOutput2.Text = "";
            txtOutput3.Text = "";
            txtOutput4.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Windows6_Load(object sender, EventArgs e)
        {

        }

        private void txtOutput1_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void Windows6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button4.PerformClick();
            }

            if (e.KeyCode == Keys.C)
            {
                button1.PerformClick();
            }

            if (e.KeyCode == Keys.R)
            {
                button2.PerformClick();
            }

            if (e.KeyCode == Keys.O)
            {
                button3.PerformClick();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求预测--节能减排1.xlsx";
            FileStream fileProcessStream1 = new FileStream(fileProcess, FileMode.Open, FileAccess.Read);
            if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook1 = new XSSFWorkbook(fileProcessStream1);  //xlsx数据读入workbook
                fileProcessStream1.Close();
            }
            else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
            {
                workbook1 = new HSSFWorkbook(fileProcessStream1);  //xls数据读入workbook
                fileProcessStream1.Close();
            }
            ISheet[] sheet1 = new ISheet[12];//创建12个表
            sheet1[0] = workbook1.GetSheetAt(0);  //获取第一个工作表  （模板中） 其实就是一个复制
            String UserName1 = txtInput1.Text;
            String UserName2 = txtOutput1.Text;
            String UserName3 = txtOutput2.Text;
            String UserName4 = txtOutput3.Text;
            String UserName5 = txtOutput4.Text;
          

            sheet1[0].GetRow(3).GetCell(3).SetCellValue(UserName1);
            sheet1[0].GetRow(5).GetCell(3).SetCellValue(UserName2);
            sheet1[0].GetRow(6).GetCell(3).SetCellValue(UserName3);
            sheet1[0].GetRow(7).GetCell(3).SetCellValue(UserName4);
            sheet1[0].GetRow(8).GetCell(3).SetCellValue(UserName5);

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析-节能减排";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
           
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }                    
            MessageBox.Show("成功导出文档至桌面");
        }
    }
    }

