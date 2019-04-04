using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows10 : Form
    {
        public Windows10()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Calculate();
        }

        #region 城市燃气项目的计算
        private void Calculate()
        {
            try
            {
                double a = 0;
                string str1 = label1.Text;
                string str2 = comboBox1.Text;
                string str3 = lblInput1.Text;
                string str4 = Txtinput1.Text;
                string str5 = lblInput2.Text;
                string str6 = Txtinput2.Text;
                string str7 = lblInput3.Text;
                string str8 = Txtinput3.Text;
                string str9 = lblInput4.Text;
                string str10 = Txtinput4.Text;
                string str11 = lblInput5.Text;
                string str12 = Txtinput5.Text;
                Common.ParameterErrorDetectionCitySize(str1, str2);
                Common.ParameterErrorDetectionCityPopulation(str3, str4);
                Common.ParameterErrorDetectionAlreadyGasUsedPopulation(str5, str6);
                Common.ParameterErrorDetectionNowResidentGasUsed(str7, str8);
                Common.ParameterErrorDetectionNowHeatingArea(str9, str10);
                Common.ParameterErrorDetectionNowYear(str11, str12);
                if (comboBox1.Text == "直辖市")
                { a = 0.95; }
                if (comboBox1.Text == "省会及计划单列市")
                { a = 0.90; }
                if (comboBox1.Text == "一般地级市")
                { a = 0.85; }
                if (comboBox1.Text == "县级市")
                { a = 0.85; }
                if (comboBox1.Text == "一般县城")
                { a = 0.80; }
                if (comboBox1.Text == "旅游城市")
                { a = 0.90; }
                //居民用气量输出计算
                int i1 = 0;
                double b = 0;
                int years = int.Parse(Txtinput5.Text);
                for (int i = years; i < years + 17; i++)
                {
                    Double p1 = Convert.ToDouble(Txtinput1.Text);//城镇人口数
                    Double p2 = Convert.ToDouble(Txtinput3.Text);//现状用气量
                    Double T1 = p1 * Math.Pow(1.02, 18) * a * 60 - p2;
                    Double T2 = T1 / 18 * (i - 2017) + p2;
                    Double residentgasuse = Math.Ceiling(T2);
                    this.dataGridView1.Rows[i1].Cells[0].Value = Convert.ToString(residentgasuse);
                    if (comboBox1.Text == "直辖市")
                    { b = 0.8; }
                    if (comboBox1.Text == "省会及计划单列市")
                    { b = 0.7; }
                    if (comboBox1.Text == "一般地级市")
                    { b = 0.6; }
                    if (comboBox1.Text == "县级市")
                    { b = 0.5; }
                    if (comboBox1.Text == "一般县城")
                    { b = 0.40; }
                    if (comboBox1.Text == "旅游城市")
                    { b = 1.2; }

                    //公福用气量输出计算

                    Double T3 = residentgasuse * b;
                    Double Publicbenefitgasuse = Math.Ceiling(T3);
                    this.dataGridView1.Rows[i1].Cells[1].Value = Convert.ToString(Publicbenefitgasuse);
                    double c = 0;
                    if (comboBox1.Text == "直辖市")
                    { c = 0.7; }
                    if (comboBox1.Text == "省会及计划单列市")
                    { c= 0.7; }
                    if (comboBox1.Text == "一般地级市")
                    { c = 0.6; }
                    if (comboBox1.Text == "县级市")
                    { c = 0.5; }
                    if (comboBox1.Text == "一般县城")
                    { c = 0.40; }
                    if (comboBox1.Text == "旅游城市")
                    { c = 0.6; }
                    //采暖用气输出计算
                    Double p3 = Convert.ToDouble(Txtinput4.Text);//现状采暖面积
                    Double p4 = Convert.ToDouble(Txtinput2.Text);//已气化人口
                    Double T4 = p1 * Math.Pow(1.02, 18) * c * 32 - p3;
                    Double T5 = (T4 / 18 * (i - 2017) + p3) * 10;
                    Double SupplyHeatingGasUse = Math.Ceiling(T5);
                    this.dataGridView1.Rows[i1].Cells[2].Value = Convert.ToString(SupplyHeatingGasUse);
                    //合计
                    Double Sum = residentgasuse + Publicbenefitgasuse + SupplyHeatingGasUse;
                    this.dataGridView1.Rows[i1].Cells[3].Value = Convert.ToString(Sum);
                    i1++;

                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        #endregion        城市
        private void Windows10_Load(object sender, EventArgs e)
        {
          
            dataGridView1.EnableHeadersVisualStyles = false;// 变灰
            this.dataGridView1.RowHeadersWidth = 70;//设置宽度
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView1.TopLeftHeaderCell.Value = "年份";
            dataGridView1.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            int years = 2035 - int.Parse(Txtinput5.Text);
            int index = this.dataGridView1.Rows.Add(years);
            int k = int.Parse(Txtinput5.Text) + 1;
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                this.dataGridView1.Rows[i].HeaderCell.Value = Convert.ToString(k);
                k++;
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            Txtinput1.Text = "";
            Txtinput2.Text = "";
            Txtinput3.Text = "";
            Txtinput4.Text = "";

            foreach (DataGridViewRow row1 in dataGridView1.Rows)
            {
                row1.Cells[0].Value = string.Empty;
                row1.Cells[1].Value = string.Empty;
                row1.Cells[2].Value = string.Empty;
                row1.Cells[3].Value = string.Empty;
            }
        }

        private void Txtinput5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Txtinput5.Text == "")
                {
                    throw new InvalidOperationException("输入参数{ 现状年份：" + Txtinput5.Text + "}为空，请重新输入。");
                }
                //int VAR = 2036 - Convert.ToInt32(Txtinput5.Text);
                //int index = this.dataGridView1.Rows.Add(VAR);
                int k = int.Parse(Txtinput5.Text);
                for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
                {
                    this.dataGridView1.Rows[i].HeaderCell.Value = Convert.ToString(k + 1);
                    k++;
                }
                foreach (DataGridViewRow row1 in dataGridView1.Rows)
                {
                    row1.Cells[0].Value = string.Empty;
                    row1.Cells[1].Value = string.Empty;
                    row1.Cells[2].Value = string.Empty;
                    row1.Cells[3].Value = string.Empty;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        private void Txtinput1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //参数检测，判断txtInput2输入是否含有字符
                foreach (char c in Txtinput1.Text)
                {
                    if (char.IsLetter(c))
                    {
                        throw new InvalidOperationException("输入参数{城镇人口：}中含有字母，请重新输入。");
                        Txtinput1.Focus();
                    }
                }

                foreach (char a in Txtinput1.Text)
                {
                    if (char.IsControl(a))
                    {
                        throw new InvalidOperationException("输入参数{城镇人口：}中含有字符，请重新输入。");
                    }
                }
                String k = Txtinput1.Text;
                if (k == "-")
                {
                    throw new InvalidOperationException("输入参数为负数，请重新输入。");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                IWorkbook workbook1 = null;  //新建IWorkbook对象
                string fileProcess = "需求分析--城市燃气1.xlsx";
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
                ISheet[] sheet1 = new ISheet[20];//创建12个表
                sheet1[0] = workbook1.GetSheetAt(0);  //获取第一个工作表  （模板中） 其实就是一个复制

                String UserName2 = comboBox1.Text;
                String UserName3 = Txtinput1.Text;
                String UserName4 = Txtinput2.Text;
                String UserName5 = Txtinput3.Text;
                String UserName6 = Txtinput4.Text;
                String UserName7 = Txtinput5.Text;

                sheet1[0].GetRow(4).GetCell(3).SetCellValue(UserName2);
                sheet1[0].GetRow(4).GetCell(5).SetCellValue(UserName3);
                sheet1[0].GetRow(4).GetCell(6).SetCellValue(UserName4);
                sheet1[0].GetRow(4).GetCell(7).SetCellValue(UserName5);
                sheet1[0].GetRow(4).GetCell(8).SetCellValue(UserName6);
                sheet1[0].GetRow(4).GetCell(9).SetCellValue(UserName7);

                #region  城市燃气datagridview的计算和导出
                double a = 0;
                if (comboBox1.Text == "直辖市")
                {
                    a = 0.95;
                }
                if (comboBox1.Text == "省会及计划单列市")
                {
                    a = 0.90;
                }
                if (comboBox1.Text == "一般地级市")
                {
                    a = 0.85;
                }
                if (comboBox1.Text == "县级市")
                {
                    a = 0.85;
                }
                if (comboBox1.Text == "一般县城")
                {
                    a = 0.80;
                }
                if (comboBox1.Text == "旅游城市")
                {
                    a = 0.90;
                }
                //居民用气量输出计算

                double b = 0;
                int years = int.Parse(Txtinput5.Text);
                int k = 4;
                for (int i = years; i < years + 18; i++)
                {
                    Double p1 = Convert.ToDouble(Txtinput1.Text);//城镇人口数
                    Double p2 = Convert.ToDouble(Txtinput3.Text);//现状用气量
                    Double T1 = p1 * Math.Pow(1.02, 18) * a * 60 - p2;
                    Double T2 = T1 / 18 * (i - 2017) + p2;
                    Double residentgasuse = Math.Ceiling(T2);
                   
                    if (comboBox1.Text == "直辖市")
                    {
                        b = 0.8;
                    }
                    if (comboBox1.Text == "省会及计划单列市")
                    {
                        b = 0.7;
                    }
                    if (comboBox1.Text == "一般地级市")
                    {
                        b = 0.6;
                    }
                    if (comboBox1.Text == "县级市")
                    {
                        b = 0.5;
                    }
                    if (comboBox1.Text == "一般县城")
                    {
                        b = 0.40;
                    }
                    if (comboBox1.Text == "旅游城市")
                    {
                        b = 1.2;
                    }

                    //公福用气量输出计算

                    Double T3 = residentgasuse * b;
                    Double Publicbenefitgasuse = Math.Ceiling(T3);
                  

                    //采暖用气输出计算
                    Double p3 = Convert.ToDouble(Txtinput4.Text);//现状采暖面积
                    Double p4 = Convert.ToDouble(Txtinput2.Text);//已气化人口
                    Double T4 = p1 * Math.Pow(1.02, 18) * 0.7 * 32 - p3;
                    Double T5 = (T4 / 18 * (i - 2017) + p3) * 10;
                    Double SupplyHeatingGasUse = Math.Ceiling(T5);
                    //合计
                    Double Sum = residentgasuse + Publicbenefitgasuse + SupplyHeatingGasUse;

                    sheet1[0].GetRow(8).GetCell(k).SetCellValue(residentgasuse.ToString("0"));
                    sheet1[0].GetRow(9).GetCell(k).SetCellValue(Publicbenefitgasuse.ToString("0"));
                    sheet1[0].GetRow(10).GetCell(k).SetCellValue(SupplyHeatingGasUse.ToString("0"));
                    sheet1[0].GetRow(11).GetCell(k).SetCellValue(Sum.ToString("0"));
                    k++;
                }
                #endregion
                string path = null;
                saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
                saveFile.FileName = "智能分析--城市燃气";//设置默认文件名
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
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Windows10_KeyDown(object sender, KeyEventArgs e)
        {
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
                button4.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button3.PerformClick();
            }
        }

        private void Windows10_Activated(object sender, EventArgs e)
        {
            this.Txtinput1.Focus();
        }
    }
}
