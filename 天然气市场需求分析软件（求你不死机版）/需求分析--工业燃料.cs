using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows9 : Form
    {
        public Windows9()
        {
            InitializeComponent();
        }
        private void ParameterErrorDetectiontextBox21()
        {

            //参数检测，判断输入是否为空
            if (txtInput1.Text == "")
            {
                throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}为空，请重新输入。");

            }
            //参数检测，判断输入是否含有字符
            foreach (char c in txtInput1.Text)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}输入参数含有字符，请重新输入。");
                }
            }

            double targetValue1 = Convert.ToDouble(txtInput1.Text);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 1000000)
            {
                throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}超过输入参数范围（0,1000000]，请重新输入。");
            }
            if (txtInput1.Text == "")
            {
                throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}为空，请重新输入。");

            }
            //参数检测，判断输入是否含有字符
            foreach (char c in txtInput1.Text)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + lblInput1.Text + txtInput1.Text + "}输入参数含有字符，请重新输入。");
                }
            }

            double targetValue2 = Convert.ToDouble(txtInput2.Text);
            //参数检测，判断输入是否为数字、零
            if (targetValue2 < 0)
            {
                throw new InvalidOperationException("输入参数{" + lblInput2.Text + txtInput2.Text + "}为负数，请重新输入。");

            }
            if (targetValue2 == 0)
            {
                throw new InvalidOperationException("输入参数{" + lblInput2.Text + txtInput2.Text + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue2 > 1000000)
            {
                throw new InvalidOperationException("输入参数{" + lblInput2.Text + txtInput2.Text + "}超过输入参数范围（0,1000000]，请重新输入。");
            }
        }
        string ProductName1;
        int[] s = new int[1000];
        int Num = 0;
        private void Form2_Load(object sender, EventArgs e)
        {

        }

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
            ProductName1 = textBox1.Text;
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
                fileProcessStream1.Close();
            }


            ISheet sheet1 = workbook1.GetSheetAt(0);
            for (int j = 0; j < Num; j++)
            {
                IRow Row = sheet.GetRow(s[j]);
                ICell cell = Row.GetCell(9);
                sheet1.GetRow(j + 4).GetCell(2).SetCellValue(sheet.GetRow(s[j]).GetCell(16).ToString());//用气需求量

                //sheet1.GetRow(j + 4).GetCell(3).SetCellValue(sheet.GetRow(s[j]).GetCell(9).ToString());//达产时间

                sheet1.GetRow(j + 4).GetCell(3).SetCellValue(cell.DateCellValue.ToString("yyyy/MM/dd"));//达产时间

                // //判断输入是否是日期
                //if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                //{
                //    sheet1.GetRow(j + 4).GetCell(3).SetCellValue(cell.DateCellValue.ToString("yyyy/MM/dd"));//达产时间

                //}
                //else
                //{
                //    sheet1.GetRow(j + 4).GetCell(3).SetCellValue(cell.ToString());//达产时间
                //}

                sheet1.GetRow(j + 4).GetCell(4).SetCellValue(sheet.GetRow(s[j]).GetCell(10).ToString());//用气压力需求（MPa）
                sheet1.GetRow(j + 4).GetCell(5).SetCellValue(sheet.GetRow(s[j]).GetCell(18).ToString());//负荷特点
                sheet1.GetRow(j + 4).GetCell(6).SetCellValue(sheet.GetRow(s[j]).GetCell(19).ToString());//高峰月
                sheet1.GetRow(j + 4).GetCell(7).SetCellValue(sheet.GetRow(s[j]).GetCell(20).ToString());//最大月不均匀系数
                sheet1.GetRow(j + 4).GetCell(8).SetCellValue(sheet.GetRow(s[j]).GetCell(21).ToString());//可承受气价（元/Nm3）
                sheet1.GetRow(j + 4).GetCell(9).SetCellValue(sheet.GetRow(s[j]).GetCell(22).ToString());//燃气成本比例（%）
            }


            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测窑炉用气产品匹配.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            fileStream.Close();
            workbook.Close();
            file.Close();
            workbook1.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void Calcuiate()
        {
            try
            {
                string str1 = lblInput1.Text;
                string str2 = txtInput1.Text;
                string str3 = lblInput2.Text;
                string str4 = txtInput2.Text;
                Common.ParameterErrorDetectionBoilerScale(str1, str2);
                Common.ParameterErrorDetectionGasTime(str3, str4);

                double Input1 = Convert.ToDouble(txtInput1.Text);  //输入量
                double Input2 = Convert.ToDouble(txtInput2.Text);  //输入量

                txtOutput1.Text = (Input1 * 70 * Input2 / 100000000).ToString("0.0000");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            Calcuiate();
        }
        private void button10_Click(object sender, EventArgs e)
        {


            txtInput1.Text = "";
            txtInput2.Text = "";
            txtOutput1.Text = "";

        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {

        }

        //private void txtInput2_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    //数字0~9所对应的keychar为48~57，小数点是46，Backspace是8
        //    e.Handled = true;
        //    //输入0-9和Backspace del 有效
        //    if ((e.KeyChar >= 47 && e.KeyChar <= 58) || e.KeyChar == 8)
        //    {
        //        e.Handled = false;
        //    }
        //    if (e.KeyChar == 46)                       //小数点      
        //    {
        //        if (txtInput2.Text.Length <= 0)
        //            e.Handled = true;           //小数点不能在第一位      
        //        else
        //        {
        //            float f;
        //            if (float.TryParse(txtInput2.Text + e.KeyChar.ToString(), out f))
        //            {
        //                e.Handled = false;
        //            }
        //        }
        //    }

        //}



        private void button5_Click(object sender, EventArgs e)
        {
            Calculate();
        }

        private void Calculate()
        {
            try
            {

                string str1 = lblInput13.Text;
                string str2 = txtInput13.Text;
                string str3 = lblInput14.Text;
                string str4 = txtInput14.Text;
                Common.ParameterErrorDetectionProductAmount(str1, str2);
                Common.ParameterErrorDetectionGasUsedAmount(str3, str4);


                double input = Convert.ToDouble(txtInput13.Text);
                double input1 = Convert.ToDouble(txtInput14.Text);
                textBox14.Text = (input * input1 / 100000000.0).ToString();

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
                textBox14.Text = "";
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }

        }

        private void buttonn2_Click(object sender, EventArgs e)
        {
            foreach (Control cc in this.groupBox2.Controls)
            {
                textBox14.Text = "";
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
            foreach (Control cc in this.groupBox3.Controls)
            {
                textBox14.Text = "";
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            foreach (Control cc in this.groupBox11.Controls)
            {
                textBox14.Text = "";
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
            foreach (Control cc in this.groupBox12.Controls)
            {
                textBox14.Text = "";
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
        }

        private void buttonn4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtInput14_TextChanged(object sender, EventArgs e)
        {

        }

        private void Windows9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button3.PerformClick();
            }
            if (e.KeyCode == Keys.C)
            {
                button5.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button6.PerformClick();
            }
            if (e.KeyCode == Keys.M)
            {
                button4.PerformClick();
            }
            if (e.KeyCode == Keys.O)
            {
                button2.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button1.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button2.PerformClick();
            }
        }

        private void tabControl1_KeyDown(object sender, KeyEventArgs e)
        {
           
            if (e.KeyCode == Keys.C)
            {
                button9.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button10.PerformClick();
            }
            if (e.KeyCode == Keys.M)
            {
                button13.PerformClick();
            }
            if (e.KeyCode == Keys.O)
            {
                button2.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button11.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button12.PerformClick();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--工业燃料--窑炉用气1.xlsx";
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
            String UserName1 = txtInput11.Text;
            String UserName2 = txtInput12.Text;
            String UserName3 = txtInput13.Text;
            String UserName4 = txtInput14.Text;
            String UserName5 = textBox14.Text;

            String UserName6 = textBox3.Text;
            String UserName7 = textBox2.Text;
            String UserName8 = textBox1.Text;
            String UserName9 = textBox13.Text;
            String UserName10 = textBox4.Text;
            String UserName11= textBox11.Text;
            String UserName12 = textBox12.Text;
            String UserName13 = textBox10.Text;
            String UserName14 = textBox8.Text;
            String UserName15 = textBox9.Text;
            String UserName16 = textBox7.Text;
            String UserName17 = textBox5.Text;
            String UserName18= textBox6.Text;
            


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
            saveFile.FileName = "需求分析--工业燃料";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--工业燃料--窑炉用气1.xlsx";
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
            String UserName1 = txtInput11.Text;
            String UserName2 = txtInput12.Text;
            String UserName3 = txtInput13.Text;
            String UserName4 = txtInput14.Text;
            String UserName5 = textBox14.Text;

            String UserName6 = textBox3.Text;
            String UserName7 = textBox2.Text;
            String UserName8 = textBox1.Text;
            String UserName9 = textBox13.Text;
            String UserName10 = textBox4.Text;
            String UserName11 = textBox11.Text;
            String UserName12 = textBox12.Text;
            String UserName13 = textBox10.Text;
            String UserName14 = textBox8.Text;
            String UserName15 = textBox9.Text;
            String UserName16 = textBox7.Text;
            String UserName17 = textBox5.Text;
            String UserName18 = textBox6.Text;



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
            saveFile.FileName = "需求分析--工业燃料";//设置默认文件名
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

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--工业燃料--锅炉用气1.xlsx";
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

            String UserName6 = textBox26.Text;
            String UserName7 = textBox25.Text;
            String UserName8 = textBox24.Text;
            String UserName9 = textBox36.Text;
            String UserName10 = textBox27.Text;
            String UserName11 = textBox34.Text;
            String UserName12 = textBox35.Text;
            String UserName13 = textBox33.Text;
            String UserName14 = textBox31.Text;
            String UserName15 = textBox32.Text;
            String UserName16 = textBox30.Text;
            String UserName17 = textBox28.Text;
            String UserName18 = textBox29.Text;



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
            saveFile.FileName = "需求分析--工业燃料--锅炉用气";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--工业燃料--锅炉用气1.xlsx";
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

            String UserName6 = textBox26.Text;
            String UserName7 = textBox25.Text;
            String UserName8 = textBox24.Text;
            String UserName9 = textBox36.Text;
            String UserName10 = textBox27.Text;
            String UserName11 = textBox34.Text;
            String UserName12 = textBox35.Text;
            String UserName13 = textBox33.Text;
            String UserName14 = textBox31.Text;
            String UserName15 = textBox32.Text;
            String UserName16 = textBox30.Text;
            String UserName17 = textBox28.Text;
            String UserName18 = textBox29.Text;



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
            saveFile.FileName = "需求分析--工业燃料--锅炉用气";//设置默认文件名
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

   

 




