using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows4 : Form
    {
        public Windows4()
        {
            InitializeComponent();
        }

        string ProductName1;
        int[] s = new int[1000];
        int Num =0;
        //double AA = 898989;
        private void button1_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象
            string fileName = "参数库.xlsx";
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
            }


            ISheet sheet = workbook.GetSheetAt(0);
           /* ProductName1 = comboBox1.Text;   */ //is have a error about input name !!
            int RowCount = sheet.LastRowNum;
            for (int i = 2; i < 83; i++)
            {

                if (ProductName1 == sheet.GetRow(i).GetCell(5).ToString())
                {
                    s[Num] = i;
                    Num++;
                    continue;

                }
            }
            //Console.ReadLine();
            //fileStream.Close();
            //workbook.Close();


            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求预测化工.xlsx";
            FileStream fileProcessStream1 = new FileStream(fileProcess, FileMode.Open, FileAccess.Read);
            if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook1 = new XSSFWorkbook(fileProcessStream1);  //xlsx数据读入workbook
                fileProcessStream1.Close();
            }
            else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
            {
                workbook1 = new HSSFWorkbook(fileProcessStream1);  //xls数据读入workbook

            }


            ISheet sheet1 = workbook1.GetSheetAt(0);
            for (int j = 0; j < Num; j++)
            {
                IRow Row = sheet.GetRow(s[j]);
                ICell cell = Row.GetCell(9);
                sheet1.GetRow(j + 4).GetCell(2).SetCellValue(sheet.GetRow(s[j]).GetCell(16).ToString());//用气需求量
                sheet1.GetRow(j + 4).GetCell(3).SetCellValue(cell.DateCellValue.ToString("yyyy/MM/dd"));//达产时间

                sheet1.GetRow(j + 4).GetCell(4).SetCellValue(sheet.GetRow(s[j]).GetCell(10).ToString());//用气压力需求（MPa）
                sheet1.GetRow(j + 4).GetCell(5).SetCellValue(sheet.GetRow(s[j]).GetCell(18).ToString());//负荷特点
                sheet1.GetRow(j + 4).GetCell(6).SetCellValue(sheet.GetRow(s[j]).GetCell(19).ToString());//高峰月
                sheet1.GetRow(j + 4).GetCell(7).SetCellValue(sheet.GetRow(s[j]).GetCell(20).ToString());//最大月不均匀系数
                sheet1.GetRow(j + 4).GetCell(8).SetCellValue(sheet.GetRow(s[j]).GetCell(21).ToString());//可承受气价（元/Nm3）
                sheet1.GetRow(j + 4).GetCell(9).SetCellValue(sheet.GetRow(s[j]).GetCell(22).ToString());//燃气成本比例（%）
            }

            //double MM;
            //MM = Convert.ToDouble(textBox1.Text);

            //sheet1.GetRow(3).GetCell(3).SetCellValue(MM);

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测化工产品匹配.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            fileStream.Close();
            workbook.Close();
            file.Close();
            workbook1.Close();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            txtOutput1.Text = "";
            foreach (Control cc in this.groupBox9.Controls)
            {
                textBox14.Text = "";
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Calculate();
        }

        private void Calculate()
        {
            try
            {

                string str1 = lblInput1.Text;
                string str2 = txtInput1.Text;
                string str3 = lblInput2.Text;
                string str4 = txtInput2.Text;
                Common.ParameterErrorDetectionProductOutput(str1, str2);
                Common.ParameterErrorDetectionGasUsed(str3, str4);

                double input = Convert.ToDouble(txtInput1.Text);
                double input1 = Convert.ToDouble(txtInput2.Text);
                txtOutput1.Text = (input * input1 / 100000000.0).ToString();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (Control cc in this.groupBox6.Controls)
            {
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
            foreach (Control cc in this.groupBox2.Controls)
            {
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Windows4_Load(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--化工1.xlsx";
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
            String UserName1 = textBox23.Text;
            String UserName2 = textBox22.Text;
            String UserName3 = txtInput1.Text;
            String UserName4 = txtInput2.Text;
            String UserName5 = txtOutput1.Text;

            String UserName6 = textBox10.Text;
            String UserName7 = textBox9.Text;
            String UserName8 = textBox17.Text;
            String UserName9 = textBox8.Text;
            String UserName10 = textBox18.Text;
            String UserName11 = textBox4.Text;
            String UserName12 = textBox7.Text;
            String UserName13 = textBox13.Text;
            String UserName14 = textBox11.Text;
            String UserName15 = textBox12.Text;
            String UserName16 = textBox16.Text;
            String UserName17 = textBox14.Text;
            String UserName18 = textBox15.Text;



            sheet1[0].GetRow(3).GetCell(2).SetCellValue(UserName1);
            sheet1[0].GetRow(4).GetCell(2).SetCellValue(UserName2);
            sheet1[0].GetRow(5).GetCell(2).SetCellValue(UserName3);
            sheet1[0].GetRow(6).GetCell(2).SetCellValue(UserName4);
            sheet1[0].GetRow(7).GetCell(2).SetCellValue(UserName5);

            sheet1[0].GetRow(3).GetCell(6).SetCellValue(UserName6);
            sheet1[0].GetRow(4).GetCell(6).SetCellValue(UserName7);
            sheet1[0].GetRow(5).GetCell(6).SetCellValue(UserName8);
            sheet1[0].GetRow(6).GetCell(6).SetCellValue(UserName9);
            sheet1[0].GetRow(7).GetCell(6).SetCellValue(UserName10);
            sheet1[0].GetRow(8).GetCell(6).SetCellValue(UserName11);
            sheet1[0].GetRow(9).GetCell(6).SetCellValue(UserName12);
            sheet1[0].GetRow(10).GetCell(6).SetCellValue(UserName13);
            sheet1[0].GetRow(11).GetCell(6).SetCellValue(UserName14);
            sheet1[0].GetRow(12).GetCell(6).SetCellValue(UserName15);
            sheet1[0].GetRow(13).GetCell(6).SetCellValue(UserName16);
            sheet1[0].GetRow(14).GetCell(6).SetCellValue(UserName17);
            sheet1[0].GetRow(15).GetCell(6).SetCellValue(UserName18);

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析--化工";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }
    
        private void button10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button8.PerformClick();
            }
            if (e.KeyCode == Keys.C)
            {
                button9.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button10.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button7.PerformClick();
            }
            if (e.KeyCode == Keys.M)
            {
                button5.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button6.PerformClick();
            }
            if (e.KeyCode == Keys.O)
            {
                button1.PerformClick();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--化工1.xlsx";
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
            String UserName1 = textBox23.Text;
            String UserName2 = textBox22.Text;
            String UserName3 = txtInput1.Text;
            String UserName4 = txtInput2.Text;
            String UserName5 = txtOutput1.Text;

            String UserName6 = textBox10.Text;
            String UserName7 = textBox9.Text;
            String UserName8 = textBox17.Text;
            String UserName9 = textBox8.Text;
            String UserName10 = textBox18.Text;
            String UserName11 = textBox4.Text;
            String UserName12 = textBox7.Text;
            String UserName13 = textBox13.Text;
            String UserName14 = textBox11.Text;
            String UserName15 = textBox12.Text;
            String UserName16 = textBox16.Text;
            String UserName17 = textBox14.Text;
            String UserName18 = textBox15.Text;



            sheet1[0].GetRow(3).GetCell(2).SetCellValue(UserName1);
            sheet1[0].GetRow(4).GetCell(2).SetCellValue(UserName2);
            sheet1[0].GetRow(5).GetCell(2).SetCellValue(UserName3);
            sheet1[0].GetRow(6).GetCell(2).SetCellValue(UserName4);
            sheet1[0].GetRow(7).GetCell(2).SetCellValue(UserName5);

            sheet1[0].GetRow(3).GetCell(6).SetCellValue(UserName6);
            sheet1[0].GetRow(4).GetCell(6).SetCellValue(UserName7);
            sheet1[0].GetRow(5).GetCell(6).SetCellValue(UserName8);
            sheet1[0].GetRow(6).GetCell(6).SetCellValue(UserName9);
            sheet1[0].GetRow(7).GetCell(6).SetCellValue(UserName10);
            sheet1[0].GetRow(8).GetCell(6).SetCellValue(UserName11);
            sheet1[0].GetRow(9).GetCell(6).SetCellValue(UserName12);
            sheet1[0].GetRow(10).GetCell(6).SetCellValue(UserName13);
            sheet1[0].GetRow(11).GetCell(6).SetCellValue(UserName14);
            sheet1[0].GetRow(12).GetCell(6).SetCellValue(UserName15);
            sheet1[0].GetRow(13).GetCell(6).SetCellValue(UserName16);
            sheet1[0].GetRow(14).GetCell(6).SetCellValue(UserName17);
            sheet1[0].GetRow(15).GetCell(6).SetCellValue(UserName18);

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析--化工";//设置默认文件名
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
