using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Collections.Generic;


namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows3 : Form

    {
        Double[] a;
        public Windows3()
        {
            InitializeComponent();
        }
        public int MiddleVar1 = 3;
        public double MiddleVar2 = 390;
        public double MiddleVar3 = 52;
        public double MiddleVar4 = 15100;
        public double MiddleVar5 = 17100;
        public double MiddleVar6 = 89;
        string ProductName1;
        int[] s = new int[1000];
        int Num = 0;


        private void ChangeLabel(int i)
        {
            if (i == 1)
            {
                label2.Text = "机组台数：";
                label5.Text = "单台容量";
                lblInput5.Text = "机组效率";

                lblInput7.Text = "台";
                label4.Text = "MW";
                lblInput6.Text = "%";
            }
            if (i == 2)
            {
                label2.Text = "年发电量：";
                label5.Text = "年供热量";
                lblInput5.Text = "年制冷量";

                lblInput7.Text = "MWh";
                label4.Text = "GJ";
                lblInput6.Text = "GJ";
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

            ChangeLabel(1);
            button1.Enabled = true;
            button2.Enabled = false;

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            ChangeLabel(2);
            button1.Enabled = false;
            button2.Enabled = true;

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

            ChangeLabel(2);
            button1.Enabled = false;
            button2.Enabled = true;

        }


        private void InstalledCapacity(int a, double b, double c, double x, double y, double z)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "工作簿166667.xlsx";
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
            sheet1.GetRow(9).GetCell(4).SetCellValue(a);
            sheet1.GetRow(10).GetCell(4).SetCellValue(b);
            sheet1.GetRow(11).GetCell(4).SetCellValue(c);
            sheet1.GetRow(9).GetCell(7).SetCellValue(x);
            sheet1.GetRow(10).GetCell(7).SetCellValue(y);
            sheet1.GetRow(11).GetCell(7).SetCellValue(z);



            for (int i = 18; i < 23; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 3; j < 10; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell.CellType == CellType.Formula)
                    {
                        IFormulaEvaluator m = null;
                        if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
                        {
                            m = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
                        {
                            m = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        m.EvaluateInCell(cell);
                    }
                }
            }

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-发电.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();
        }

        private void ProductOutput(double x, double y, double z)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "工作簿166667.xlsx";
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
            sheet1.GetRow(9).GetCell(7).SetCellValue(x);
            sheet1.GetRow(10).GetCell(7).SetCellValue(y);
            sheet1.GetRow(11).GetCell(7).SetCellValue(z);



            for (int i = 18; i < 23; i++)
            {
                IRow row = sheet1.GetRow(i);
                for (int j = 3; j < 10; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell.CellType == CellType.Formula)
                    {
                        IFormulaEvaluator m = null;
                        if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
                        {
                            m = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
                        {
                            m = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        }
                        m.EvaluateInCell(cell);
                    }
                }
            }

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-发电.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();

        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

            ChangeLabel(1);
            button1.Enabled = true;
            button2.Enabled = false;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            MiddleVar1 = Convert.ToInt32(txtinput17.Text);
            MiddleVar2 = Convert.ToDouble(txtinput16.Text);
            MiddleVar3 = Convert.ToDouble(txtinput15.Text);

            InstalledCapacity(MiddleVar1, MiddleVar2, MiddleVar3, MiddleVar4, MiddleVar5, MiddleVar6);


        }

        private void button2_Click(object sender, EventArgs e)
        {

            MiddleVar4 = Convert.ToDouble(txtinput17.Text);
            MiddleVar5 = Convert.ToDouble(txtinput16.Text);
            MiddleVar6 = Convert.ToDouble(txtinput15.Text);

            InstalledCapacity(MiddleVar1, MiddleVar2, MiddleVar3, MiddleVar4, MiddleVar5, MiddleVar6);

        }

        private void button3_Click(object sender, EventArgs e)
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
            ProductName1 = txtinput13.Text;
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


            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测化工产品匹配567890.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            fileStream.Close();
            workbook.Close();
            file.Close();
            workbook1.Close();
        }



        private void Windows3_Load(object sender, EventArgs e)
        {
            dataGridView1.EnableHeadersVisualStyles = false;// 变灰
            dataGridView2.EnableHeadersVisualStyles = false;// 变灰



            int index1 = this.dataGridView1.Rows.Add(8);
            dataGridView1.TopLeftHeaderCell.Value = "序号";
            this.dataGridView1.RowHeadersWidth = 50;//设置宽度
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView1.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dataGridView2.RowHeadersWidth = 50;//设置宽度
            dataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView2.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            int k = 1;
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView1.Rows[i].HeaderCell.Value = Convert.ToString(k);
                k++;
            }

            int index = this.dataGridView2.Rows.Add(8);
            dataGridView2.TopLeftHeaderCell.Value = "序号";
            int j = 1;
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView2.Rows[i].HeaderCell.Value = Convert.ToString(j);
                j++;
            }

        }
        private double easy(double p)
        {
            Double p1 = Convert.ToDouble(txtinput13.Text);//机组台数
            Double p2 = Convert.ToDouble(txtinput14.Text);//单台容量

            Double p11 = Convert.ToDouble(txtinput17.Text);//天然气热值
            Double p12 = Convert.ToDouble(txtinput15.Text);//机组效率

            double MiddleVar = p1 * p2 * 1000000 * p * 3600 / (p11 * 1000 * 4.18 * p12 * 1000000);
            return MiddleVar;
        }
        private void Calculate()
        {
            try
            {
                string str1 = lblInput3.Text;
                string str2 = txtinput13.Text;
                string str3 = lblInput4.Text;
                string str4 = txtinput14.Text;
                string str5 = lblInput5.Text;
                string str6 = txtinput15.Text;
                string str9 = lblInput7.Text;
                string str10 = txtinput17.Text;
                Common.ParameterErrorDetectionMachineAmount(str1, str2);
                Common.ParameterErrorDetectionSingleMachineValume(str3, str4);
                Common.ParameterErrorDetectionMachineEfficiency(str5, str6);
                Common.ParameterErrorDetectionGasHeating(str9, str10);
                string str = txtinput16.Text;
                string[] strName = str.Split(',');
                double m = strName.Length;
                if (m > 8)
                {
                    //找出数组中数据个数
                    int num = 0;
                    for (int i = 0; i < strName.Length; i++)
                    {
                        num += 1;
                    }
                    int index = this.dataGridView1.Rows.Add(num - 8);
                }
                int k = 1;
                int j = 0;
                for (int i = 0; i < m; i++)
                {
                    this.dataGridView1.Rows[i].HeaderCell.Value = Convert.ToString(k);
                    this.dataGridView1.Rows[i].Cells[0].Value = strName[j];
                    Double T = Convert.ToDouble(strName[j]);
                    this.dataGridView1.Rows[i].Cells[1].Value = easy(T).ToString("0.0");
                    j++;
                    k++;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            TableClear();
            foreach (DataGridViewRow row1 in dataGridView1.Rows)
            {
                row1.Cells[0].Value = string.Empty;
                row1.Cells[1].Value = string.Empty;
            }
            txtinput16.Text = "";
        }

        private void TableClear()
        {
            foreach (Control cc in this.groupBox3.Controls)
            {
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }

            foreach (DataGridViewRow row1 in dataGridView1.Rows)
            {
                row1.Cells[0].Value = string.Empty;
                row1.Cells[1].Value = string.Empty;
            }
        }


        private void txtClear(Control ctls)
        {


            foreach (Control tb in ctls.Controls)
            {
                tb.Text = "";
            }
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            Calculate();
            dataGridView1.AllowUserToAddRows = false;
        }
        private void button9_Click(object sender, EventArgs e)
        {
            Calculate2();
        }
        private double Easy(double n)
        {
            Double o11 = Convert.ToDouble(txtinput27.Text);//年发电量
            Double o12 = Convert.ToDouble(txtinput28.Text);//年供热量

            Double o13 = Convert.ToDouble(txtinput29.Text);//年制冷量
            Double o14 = Convert.ToDouble(txtinput30.Text);//天然气热值
            double MiddleVar2 = (o11 * 1000 * 3600 + (o12 + o13) * 1000000) / (o14 * 4.18 * n * 1000000);
            return MiddleVar2;
        }
        private void Calculate2()
        {
            try
            {
                string str11 = lblInput23.Text;
                string str12 = txtinput23.Text;
                string str13 = lblInput24.Text;
                string str14 = txtinput24.Text;
                string str15 = lblInput25.Text;
                string str16 = txtinput25.Text;
                string str17 = lblInput26.Text;
                string str18 = txtinput26.Text;
                string str19 = lblInput27.Text;
                string str20 = txtinput27.Text;
                string str21 = lblInput28.Text;
                string str22 = txtinput28.Text;
                string str23 = lblInput29.Text;
                string str24 = txtinput29.Text;
                string str25 = lblInput30.Text;
                string str26 = txtinput30.Text;
                Common.ParameterErrorDetectionMachineAmount2(str11, str12);
                Common.ParameterErrorDetectionSingleMachineValume2(str13, str14);
                Common.ParameterErrorDetectionMachineEfficiency2(str15, str16);
                Common.ParameterErrorDetectionAnnualUsedhours(str17, str18);
                Common.ParameterErrorDetectionAnnualElectrityProduct(str19, str20);
                Common.ParameterErrorDetectionAnnualHeatingProduct(str21, str22);
                Common.ParameterErrorDetectionAnnualColdProduct(str23, str24);
                Common.ParameterErrorDetectionNaturalGasHeatValue(str25, str26);
                string str4 = txtinput25.Text;
                string[] str = str4.Split(',');
                double L = str.Length;
                if (L > 8)
                {
                    //找出数组中数据个数
                    int num1 = 0;
                    for (int i = 0; i < str.Length; i++)
                    {
                        num1 += 1;
                    }
                    int index = this.dataGridView2.Rows.Add(num1 - 8);
                }
                int k1 = 1;
                int j1 = 0;
                for (int i = 0; i < L; i++)
                {
                    this.dataGridView2.Rows[i].HeaderCell.Value = Convert.ToString(k1);
                    //this.dataGridView2.Rows[i].HeaderCell.Value = DataGridViewContentAlignment.MiddleCenter;
                    this.dataGridView2.Rows[i].Cells[0].Value = str[j1];
                    Double T = Convert.ToDouble(str[j1]);
                    this.dataGridView2.Rows[i].Cells[1].Value = Easy(T).ToString("0.0");
                    j1++;
                    k1++;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            CLearAll();
        }

        private void CLearAll()
        {
            foreach (Control cc in this.groupBox9.Controls)
            {
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
            foreach (DataGridViewRow row1 in dataGridView2.Rows)
            {
                row1.Cells[0].Value = string.Empty;
                row1.Cells[1].Value = string.Empty;
            }
        }
        private void txtinput25_TextChanged(object sender, EventArgs e)
        {
            TableClear1();
        }

        private void TableClear1()
        {
            foreach (DataGridViewRow row1 in dataGridView2.Rows)
            {
                row1.Cells[0].Value = string.Empty;
                row1.Cells[1].Value = string.Empty;

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
            foreach (Control cc in this.groupBox2.Controls)
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

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }
        private void tabControl1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button7.PerformClick();
            }

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
            if (e.KeyCode == Keys.M)
            {
                button13.PerformClick();
            }
        }

        private void Windows3_KeyDown(object sender, KeyEventArgs e)
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
            if (e.KeyCode == Keys.M)
            {
                button5.PerformClick();
            }
        }
        #region  发电调峰/基荷电厂导出至Excel
        private void button4_Click(object sender, EventArgs e)
        {

            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--调峰、基荷电厂1.xlsx";
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
            String UserName1 = textBox7.Text;
            String UserName2 = comboBox1.Text;
            String UserName3 = txtinput13.Text;
            String UserName4 = txtinput14.Text;
            String UserName5 = txtinput15.Text;
            String UserName6 = txtinput16.Text;
            String UserName7 = txtinput17.Text;
            String UserName13 = textBox10.Text;
            String UserName14 = textBox9.Text;
            String UserName15 = textBox17.Text;
            String UserName16 = textBox38.Text;
            String UserName17 = textBox18.Text;
            String UserName18 = textBox8.Text;
            String UserName19 = textBox37.Text;
            String UserName20 = textBox13.Text;
            String UserName21 = textBox11.Text;
            String UserName22 = textBox16.Text;
            String UserName23 = textBox14.Text;
            String UserName24 = textBox15.Text;
            sheet1[0].GetRow(5).GetCell(2).SetCellValue(UserName1);
            sheet1[0].GetRow(6).GetCell(2).SetCellValue(UserName2);
            sheet1[0].GetRow(7).GetCell(2).SetCellValue(UserName3);
            sheet1[0].GetRow(8).GetCell(2).SetCellValue(UserName4);
            sheet1[0].GetRow(9).GetCell(2).SetCellValue(UserName5);
            sheet1[0].GetRow(10).GetCell(2).SetCellValue(UserName6);
            sheet1[0].GetRow(11).GetCell(2).SetCellValue(UserName7);
            sheet1[0].GetRow(4).GetCell(6).SetCellValue(UserName13);
            sheet1[0].GetRow(5).GetCell(6).SetCellValue(UserName14);
            sheet1[0].GetRow(6).GetCell(6).SetCellValue(UserName15);
            sheet1[0].GetRow(7).GetCell(6).SetCellValue(UserName16);
            sheet1[0].GetRow(8).GetCell(6).SetCellValue(UserName17);
            sheet1[0].GetRow(9).GetCell(6).SetCellValue(UserName18);
            sheet1[0].GetRow(10).GetCell(6).SetCellValue(UserName19);
            sheet1[0].GetRow(11).GetCell(6).SetCellValue(UserName20);
            sheet1[0].GetRow(12).GetCell(6).SetCellValue(UserName21);
            sheet1[0].GetRow(13).GetCell(6).SetCellValue(UserName22);
            sheet1[0].GetRow(14).GetCell(6).SetCellValue(UserName23);
            sheet1[0].GetRow(15).GetCell(6).SetCellValue(UserName24);


            string str = txtinput16.Text;
            string[] strName = str.Split(',');
            double m = strName.Length;
            if (m > 8)
            {
                //找出数组中数据个数
                int num = 0;
                for (int i = 0; i < strName.Length; i++)
                {
                    num += 1;
                }
                int index = this.dataGridView1.Rows.Add(num - 8);
            }
            int j = 0;
            for (int i = 0; i < m; i++)
            {
                sheet1[0].GetRow(14 + i).GetCell(1).SetCellValue(strName[j]);
                Double T = Convert.ToDouble(strName[j]);
                sheet1[0].GetRow(14 + i).GetCell(2).SetCellValue(easy(T).ToString("0.0"));
                j++;
            }
            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析--发电";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }
        #endregion
        #region  需求分析--热电联产、三联供导出至  excel
        private void button8_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--热电联产、三联供1.xlsx";
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
            ISheet[] sheet1 = new ISheet[20];      //创建12个表
            sheet1[0] = workbook1.GetSheetAt(0);  //获取第一个工作表  （模板中） 其实就是一个复制

            String UserName1 = textBox31.Text;
            String UserName2 = comboBox2.Text;
            String UserName3 = txtinput23.Text;
            String UserName4 = txtinput24.Text;
            String UserName5 = txtinput25.Text;
            String UserName6 = txtinput26.Text;
            String UserName7 = txtinput27.Text;
            String UserName8 = txtinput28.Text;
            String UserName9 = txtinput29.Text;
            String UserName10 = txtinput30.Text;

            String UserName11 = textBox34.Text;
            String UserName12 = textBox33.Text;
            String UserName13 = textBox32.Text;
            String UserName14 = textBox46.Text;
            String UserName15 = textBox35.Text;
            String UserName16 = textBox44.Text;
            String UserName17 = textBox45.Text;
            String UserName18 = textBox43.Text;
            String UserName19 = textBox41.Text;
            String UserName20 = textBox40.Text;
            String UserName21 = textBox36.Text;
            String UserName22 = textBox39.Text;

            sheet1[0].GetRow(4).GetCell(2).SetCellValue(UserName1);
            sheet1[0].GetRow(5).GetCell(2).SetCellValue(UserName2);
            sheet1[0].GetRow(6).GetCell(2).SetCellValue(UserName3);
            sheet1[0].GetRow(7).GetCell(2).SetCellValue(UserName4);
            sheet1[0].GetRow(8).GetCell(2).SetCellValue(UserName5);
            sheet1[0].GetRow(9).GetCell(2).SetCellValue(UserName6);
            sheet1[0].GetRow(10).GetCell(2).SetCellValue(UserName7);
            sheet1[0].GetRow(11).GetCell(2).SetCellValue(UserName8);
            sheet1[0].GetRow(12).GetCell(2).SetCellValue(UserName9);
            sheet1[0].GetRow(13).GetCell(2).SetCellValue(UserName10);

            sheet1[0].GetRow(3).GetCell(5).SetCellValue(UserName11);
            sheet1[0].GetRow(4).GetCell(5).SetCellValue(UserName12);
            sheet1[0].GetRow(5).GetCell(5).SetCellValue(UserName13);
            sheet1[0].GetRow(6).GetCell(5).SetCellValue(UserName14);
            sheet1[0].GetRow(7).GetCell(5).SetCellValue(UserName15);
            sheet1[0].GetRow(8).GetCell(5).SetCellValue(UserName16);
            sheet1[0].GetRow(9).GetCell(5).SetCellValue(UserName17);
            sheet1[0].GetRow(10).GetCell(5).SetCellValue(UserName18);
            sheet1[0].GetRow(11).GetCell(5).SetCellValue(UserName19);
            sheet1[0].GetRow(12).GetCell(5).SetCellValue(UserName20);
            sheet1[0].GetRow(13).GetCell(5).SetCellValue(UserName21);
            sheet1[0].GetRow(14).GetCell(5).SetCellValue(UserName22);

            string str = txtinput25.Text;
            string[] strName = str.Split(',');
            double m = strName.Length;
            if (m > 8)
            {
                //找出数组中数据个数
                int num = 0;
                for (int i = 0; i < strName.Length; i++)
                {
                    num += 1;
                }
                int index = this.dataGridView1.Rows.Add(num - 8);
            }
            int j = 0;
            for (int i = 0; i < m; i++)
            {
                sheet1[0].GetRow(16 + i).GetCell(1).SetCellValue(strName[j]);
                Double T = Convert.ToDouble(strName[j]);
                sheet1[0].GetRow(16 + i).GetCell(2).SetCellValue(Easy(T).ToString("0.0"));
                j++;
            }
            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析--热电联产、三联供";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }
        #endregion
        #region  发电调峰/基荷电厂第二个导出
        private void button7_Click(object sender, EventArgs e)
        {

            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--调峰、基荷电厂1.xlsx";
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
            String UserName1 = textBox7.Text;
            String UserName2 = comboBox1.Text;
            String UserName3 = txtinput13.Text;
            String UserName4 = txtinput14.Text;
            String UserName5 = txtinput15.Text;
            String UserName6 = txtinput16.Text;
            String UserName7 = txtinput17.Text;
            String UserName13 = textBox10.Text;
            String UserName14 = textBox9.Text;
            String UserName15 = textBox17.Text;
            String UserName16 = textBox38.Text;
            String UserName17 = textBox18.Text;
            String UserName18 = textBox8.Text;
            String UserName19 = textBox37.Text;
            String UserName20 = textBox13.Text;
            String UserName21 = textBox11.Text;
            String UserName22 = textBox16.Text;
            String UserName23 = textBox14.Text;
            String UserName24 = textBox15.Text;

            sheet1[0].GetRow(5).GetCell(2).SetCellValue(UserName1);
            sheet1[0].GetRow(6).GetCell(2).SetCellValue(UserName2);
            sheet1[0].GetRow(7).GetCell(2).SetCellValue(UserName3);
            sheet1[0].GetRow(8).GetCell(2).SetCellValue(UserName4);
            sheet1[0].GetRow(9).GetCell(2).SetCellValue(UserName5);
            sheet1[0].GetRow(10).GetCell(2).SetCellValue(UserName6);
            sheet1[0].GetRow(11).GetCell(2).SetCellValue(UserName7);

            sheet1[0].GetRow(4).GetCell(6).SetCellValue(UserName13);
            sheet1[0].GetRow(5).GetCell(6).SetCellValue(UserName14);
            sheet1[0].GetRow(6).GetCell(6).SetCellValue(UserName15);
            sheet1[0].GetRow(7).GetCell(6).SetCellValue(UserName16);
            sheet1[0].GetRow(8).GetCell(6).SetCellValue(UserName17);
            sheet1[0].GetRow(9).GetCell(6).SetCellValue(UserName18);
            sheet1[0].GetRow(10).GetCell(6).SetCellValue(UserName19);
            sheet1[0].GetRow(11).GetCell(6).SetCellValue(UserName20);
            sheet1[0].GetRow(12).GetCell(6).SetCellValue(UserName21);
            sheet1[0].GetRow(13).GetCell(6).SetCellValue(UserName22);
            sheet1[0].GetRow(14).GetCell(6).SetCellValue(UserName23);
            sheet1[0].GetRow(15).GetCell(6).SetCellValue(UserName24);

            string str = txtinput16.Text;
            string[] strName = str.Split(',');
            double m = strName.Length;
            if (m > 8)
            {
                //找出数组中数据个数
                int num = 0;
                for (int i = 0; i < strName.Length; i++)
                {
                    num += 1;
                }
                int index = this.dataGridView1.Rows.Add(num - 8);
            }
            int j = 0;
            for (int i = 0; i < m; i++)
            {
                sheet1[0].GetRow(14 + i).GetCell(1).SetCellValue(strName[j]);
                Double T = Convert.ToDouble(strName[j]);
                sheet1[0].GetRow(14 + i).GetCell(2).SetCellValue(easy(T).ToString("0.0"));
                j++;
            }
            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析--发电";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }
        #endregion

        private void button11_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--热电联产、三联供1.xlsx";
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
            ISheet[] sheet1 = new ISheet[20];      //创建12个表
            sheet1[0] = workbook1.GetSheetAt(0);  //获取第一个工作表  （模板中） 其实就是一个复制

            String UserName1 = textBox31.Text;
            String UserName2 = comboBox2.Text;
            String UserName3 = txtinput23.Text;
            String UserName4 = txtinput24.Text;
            String UserName5 = txtinput25.Text;
            String UserName6 = txtinput26.Text;
            String UserName7 = txtinput27.Text;
            String UserName8 = txtinput28.Text;
            String UserName9 = txtinput29.Text;
            String UserName10 = txtinput30.Text;

            String UserName11 = textBox34.Text;
            String UserName12 = textBox33.Text;
            String UserName13 = textBox32.Text;
            String UserName14 = textBox46.Text;
            String UserName15 = textBox35.Text;
            String UserName16 = textBox44.Text;
            String UserName17 = textBox45.Text;
            String UserName18 = textBox43.Text;
            String UserName19 = textBox41.Text;
            String UserName20 = textBox40.Text;
            String UserName21 = textBox36.Text;
            String UserName22 = textBox39.Text;

            sheet1[0].GetRow(4).GetCell(2).SetCellValue(UserName1);
            sheet1[0].GetRow(5).GetCell(2).SetCellValue(UserName2);
            sheet1[0].GetRow(6).GetCell(2).SetCellValue(UserName3);
            sheet1[0].GetRow(7).GetCell(2).SetCellValue(UserName4);
            sheet1[0].GetRow(8).GetCell(2).SetCellValue(UserName5);
            sheet1[0].GetRow(9).GetCell(2).SetCellValue(UserName6);
            sheet1[0].GetRow(10).GetCell(2).SetCellValue(UserName7);
            sheet1[0].GetRow(11).GetCell(2).SetCellValue(UserName8);
            sheet1[0].GetRow(12).GetCell(2).SetCellValue(UserName9);
            sheet1[0].GetRow(13).GetCell(2).SetCellValue(UserName10);

            sheet1[0].GetRow(3).GetCell(5).SetCellValue(UserName11);
            sheet1[0].GetRow(4).GetCell(5).SetCellValue(UserName12);
            sheet1[0].GetRow(5).GetCell(5).SetCellValue(UserName13);
            sheet1[0].GetRow(6).GetCell(5).SetCellValue(UserName14);
            sheet1[0].GetRow(7).GetCell(5).SetCellValue(UserName15);
            sheet1[0].GetRow(8).GetCell(5).SetCellValue(UserName16);
            sheet1[0].GetRow(9).GetCell(5).SetCellValue(UserName17);
            sheet1[0].GetRow(10).GetCell(5).SetCellValue(UserName18);
            sheet1[0].GetRow(11).GetCell(5).SetCellValue(UserName19);
            sheet1[0].GetRow(12).GetCell(5).SetCellValue(UserName20);
            sheet1[0].GetRow(13).GetCell(5).SetCellValue(UserName21);
            sheet1[0].GetRow(14).GetCell(5).SetCellValue(UserName22);

            string str = txtinput25.Text;
            string[] strName = str.Split(',');
            double m = strName.Length;
            if (m > 8)
            {
                //找出数组中数据个数
                int num = 0;
                for (int i = 0; i < strName.Length; i++)
                {
                    num += 1;
                }
                int index = this.dataGridView1.Rows.Add(num - 8);
            }
            int j = 0;
            for (int i = 0; i < m; i++)
            {
                sheet1[0].GetRow(16 + i).GetCell(1).SetCellValue(strName[j]);
                Double T = Convert.ToDouble(strName[j]);
                sheet1[0].GetRow(16 + i).GetCell(2).SetCellValue(Easy(T).ToString("0.0"));
                j++;
            }
            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "需求分析--热电联产、三联供";//设置默认文件名
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                path = saveFile.FileName;
                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                workbook1.Write(file);
                file.Close();
                workbook1.Close();
            }
        }
    }
}

