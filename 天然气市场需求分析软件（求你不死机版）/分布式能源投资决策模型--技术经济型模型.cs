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
    public partial class DistributeEnergy1 : Form
    {
        public DistributeEnergy1()
        {
            InitializeComponent();
        }
        CalculateMath Calculatemath = new CalculateMath();
        private void ClearText()
        {
            foreach (Control cc in this.groupBox3.Controls)
            {
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }
            foreach (Control cc in this.groupBox4.Controls)
            {
                if (cc is TextBox)
                {
                    //清掉含有TexBox控件上的内容
                    cc.Text = "";
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ClearText1();
        }

        private void ClearText1()

        {
            foreach (DataGridViewRow row1 in dataGridView1.Rows)
            {
                for (int i = 0; i < 9; i++)
                {
                    row1.Cells[i].Value = string.Empty;
                }
            }
            foreach (Control cc in this.groupBox1.Controls)
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
                foreach (Control c in this.groupBox3.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox4.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox5.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox6.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
            }
        }

        private void 分布式能源_Load(object sender, EventArgs e)
        {
            dataGridView1.EnableHeadersVisualStyles = false;// 变灰
            dataGridView1.TopLeftHeaderCell.Value = "第N年";
            int index = this.dataGridView1.Rows.Add(32);
            this.dataGridView1.RowHeadersWidth = 61;//设置宽度
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView1.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Double k = 1;
            for (int i = 0; i < 31; i++)
            {
                this.dataGridView1.Rows[i].HeaderCell.Value = Convert.ToString(k);
                k++;
            }
            this.dataGridView1.Rows[31].HeaderCell.Value = Convert.ToString("合计");

            dataGridView2.EnableHeadersVisualStyles = false;// 变灰
            dataGridView2.TopLeftHeaderCell.Value = "第N年";
            int index1 = this.dataGridView2.Rows.Add(32);
            this.dataGridView2.RowHeadersWidth = 60;//设置宽度
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView2.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Double k1 = 1;
            for (int i = 0; i < 31; i++)
            {
                this.dataGridView2.Rows[i].HeaderCell.Value = Convert.ToString(k1);
                k1++;
            }
            this.dataGridView2.Rows[31].HeaderCell.Value = Convert.ToString("合计");



        }




        #region  收入/成本计算
        private void CostCalculate()
        {
            try
            {
                Double P8 = Convert.ToDouble(txtInput10.Text);//年利用小时数
                Double P9 = Convert.ToDouble(txtInput11.Text) / 100;//负荷率
                double Var100 = Convert.ToDouble(txtInput7.Text) * P8 / 10000 * P9; //发电量
                Double P10 = Convert.ToDouble(txtInput16.Text);//单位电价
                double Var101 = P10 * Var100;//1-电费
                Double P11 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量
                double Var102 = P11 / 10000 * P8 / 4 * P9;//制热量             
                Double P12 = Convert.ToDouble(txtInput17.Text);//单位热价（含税）
                double Var103 = P12 / (Math.Pow(10, 9) / 4182 / 860);//单位热价              
                double Var104 = Var102 * Var103;//2-热费
                Double P13 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量
                double Var105 = P13 / 10000 * P8 / 4 * 3 * P9;//制冷量               
                Double P14 = Convert.ToDouble(txtInput18.Text);//单位冷价（含税）
                double Var106 = P14 / (Math.Pow(10, 9) / 4182 / 860);//单位冷价               
                double Var107 = Var106 * Var105;//3-冷费               
                double Var108 = Var101 + Var104 + Var107;//总收入

                double Var109 = Convert.ToDouble(txtInput19.Text);//单位成本               
                Double P15 = Convert.ToDouble(txtInput14.Text);//发电气耗
                double Var110 = P15 * P9 * Var100;//用气量   
                double Var111 = Var109 * Var110;//1-燃料成本
                double Var112 = Convert.ToDouble(txtInput20.Text);//折旧年限              
                Double P1 = Convert.ToDouble(txtInput7.Text); //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资
                double Var113 = P1 * P2 / 10000;//1-总投资
                double Var114 = Var113 / Var112;//2-折旧
                Double P16 = Convert.ToDouble(txtOutput27.Text);//贷款比例
                Double P17 = Convert.ToDouble(txtOutput28.Text);//贷款利率
                Double P3 = Convert.ToDouble(txtOutput3.Text);//单位补贴
                double Var115 = P3 * P1 / 10000;//补贴
                double Var116 = Var113 - Var115; //3 - 总投资(去除补贴）
                double Var117 = Var116 * P16 * P17;//3-财务成本               
                double Var118 = Convert.ToDouble(txtInput22.Text);//人员数
                double Var123 = Convert.ToDouble(txtOutput31.Text);//人员费用
                double Var119 = Var118 * Var123;//4-人工成本
                double Var120 = Convert.ToDouble(txtInput21.Text);//单位运维成本
                double Var121 = Var100 * Var120;//5-运维等成本                                                      
                double Var122 = Var111 + Var114 + Var117 + Var119 + Var121; //总成本


                txtOutput11.Text = (Math.Floor(Var101)).ToString();//1-电费
                txtOutput12.Text = (Var100).ToString(); //发电量
                txtOutput13.Text = (P10).ToString();//单位电价
                txtOutput14.Text = (Math.Ceiling(Var104)).ToString();//2-热费
                txtOutput15.Text = (Var102).ToString();//制热量
                txtOutput16.Text = (Var103).ToString("0.00");//单位热价
                txtOutput17.Text = (Var107).ToString();//3-冷费
                txtOutput18.Text = (Var105).ToString();//制冷量
                txtOutput19.Text = (Var106).ToString("0.00");//单位冷价
                txtOutput20.Text = Math.Ceiling(Var108).ToString();  //总收入

                txtOutput21.Text = Math.Ceiling(Var111).ToString();//1-燃料成本
                txtOutput22.Text = Var109.ToString();//单位成本
                txtOutput23.Text = (Var110).ToString();//用气量
                txtOutput24.Text = Math.Ceiling(Var114).ToString();//2-折旧
                txtOutput25.Text = Var112.ToString();//折旧年限
                txtOutput26.Text = Math.Ceiling(Var117).ToString();//3-财务成本
                txtOutput29.Text = (Var119).ToString();//4-人工成本
                txtOutput30.Text = Var118.ToString();//人员数
                txtOutput32.Text = (Var121).ToString();//5-运维等成本
                txtOutput33.Text = Var120.ToString();//单位运维成本
                txtOutput34.Text = Math.Ceiling(Var122).ToString();   //总成本

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region  投资/效益计算
        private void InvestCalculate()
        {
            try
            {
                Double P1 = Convert.ToDouble(txtInput7.Text); //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资
                txtOutput1.Text = (P1 * P2 / 10000).ToString("0");//1-总投资
                Double P3 = Convert.ToDouble(txtOutput3.Text);//单位补贴
                txtOutput2.Text = (P3 * P1 / 10000).ToString();//补贴
                txtOutput4.Text = (Convert.ToDouble(txtOutput1.Text) - Convert.ToDouble(txtOutput2.Text)).ToString("0"); //总投资（去除补贴）

                Double M1 = Convert.ToDouble(txtInput10.Text);
                Double M2 = Convert.ToDouble(txtInput11.Text);
                Double M3 = Convert.ToDouble(txtInput7.Text);
                Double M4 = Convert.ToDouble(txtInput16.Text);
                Double M5 = Convert.ToDouble(txtInput8.Text);
                Double M6 = Convert.ToDouble(txtInput17.Text);
                Double M7 = Convert.ToDouble(txtInput18.Text);
                Double M19 = Convert.ToDouble(txtInput9.Text);
                Double P4 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);
                //txtInput14  txtInput11 txtInput10  txtInput7 txtInput19  txtInput20  txtInput12  txtOutput27  txtOutput28 txtOutput3 txtInput22  txtOutput31 txtInput21
                Double M8 = Convert.ToDouble(txtInput14.Text);
                Double M10 = Convert.ToDouble(txtInput19.Text);
                Double M11 = Convert.ToDouble(txtInput20.Text);
                Double M12 = Convert.ToDouble(txtInput12.Text);
                Double M13 = Convert.ToDouble(txtOutput27.Text);
                Double M14 = Convert.ToDouble(txtOutput28.Text);
                Double M15 = Convert.ToDouble(txtOutput3.Text);
                Double M16 = Convert.ToDouble(txtInput22.Text);
                Double M17 = Convert.ToDouble(txtOutput31.Text);
                Double M18 = Convert.ToDouble(txtInput21.Text);


                //Double P5 = Convert.ToDouble(txtOutput34.Text);//总成本
                Double P5 = Calculatemath.AllCost(M8, M2, M1, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);
                txtOutput5.Text = (P4 - P5).ToString("0"); //利润总额
                Double M20 = Convert.ToDouble(txtOutput5.Text);
                //2-所得税
                if (M20 > 0)
                {
                    Double P7 = Convert.ToDouble(txtOutput7.Text) / 100;
                    txtOutput6.Text = (M20 * P7).ToString("0");
                }
                //3-净利润
                txtOutput8.Text = (M20 - Convert.ToDouble(txtOutput6.Text)).ToString("0");
                //4-IRR
                //5-财务净现值  没有计算

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        #endregion

        
        private void DistributeEnergy1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.C)
            {
                button3.PerformClick();
            }
            if (e.KeyCode == Keys.R)
            {
                button1.PerformClick();
            }

            if (e.KeyCode == Keys.O)
            {
                button2.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button18.PerformClick();
            }
        }
        private void tabControl2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button4.PerformClick();
            }
        }

        private void tabControl1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button4.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button9.PerformClick();
            }
            if (e.KeyCode == Keys.E)
            {
                button16.PerformClick();
            }
        }
        #region     1-电价盈亏平衡点    
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox10.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";            
                
                #endregion
                
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion

                #region  DateGridView表中数据的计算
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
                double Var1 = 0.57;
                txtInput16.Text = Var1.ToString();                                   //单位电价（含税）
                textBox6.Text= Var1.ToString();
                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M2 = Convert.ToDouble(txtInput11.Text); //负荷率：
                Double M3 = Convert.ToDouble(txtInput7.Text);// 发电机组装机容量：
                Double M4 = Convert.ToDouble(txtInput16.Text);//单位电价：
                Double M5 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量：
                Double M6 = Convert.ToDouble(txtInput17.Text);//单位热价：
                Double M7 = Convert.ToDouble(txtInput18.Text);//单位冷价：
                Double M19 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量：
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);
                Double M9 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M10 = Convert.ToDouble(txtInput19.Text);//单位气价：
                Double M11 = Convert.ToDouble(txtInput20.Text);//折旧年限：
                Double M12 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资：
                Double M13 = Convert.ToDouble(txtOutput27.Text);//贷款比例：
                Double M14 = Convert.ToDouble(txtOutput28.Text); //贷款利率：
                Double M15 = Convert.ToDouble(txtOutput3.Text);//单位补贴：
                Double M16 = Convert.ToDouble(txtInput22.Text);//站场定员：
                Double M17 = Convert.ToDouble(txtOutput31.Text);// 人员费用：
                Double M18 = Convert.ToDouble(txtInput21.Text);// 单位运维成本：
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
                InvestCalculate();
                CostCalculate();
                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 

                this.dataGridView2.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView2.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView2.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 

                #region 填充为零
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                #endregion

                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #endregion
                }
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] =Convert.ToString(0) ;
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] =( Convert.ToDouble(a[i])+ Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8+i] = (Convert.ToDouble(a[i+8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                #endregion
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region  恢复默认设置
        private void button17_Click_1(object sender, EventArgs e)
        {
            try
            {
                txtInput7.Text = "87000";
                txtInput8.Text = "28000";
                txtInput9.Text = "0";
                txtInput10.Text = "5500";
                txtInput11.Text = "100";
                txtInput12.Text = "5686";
                txtInput13.Text = "0";
                txtInput14.Text = "0.16";
                txtInput15.Text = "80";
                txtInput16.Text = "0.7493";
                txtInput17.Text = "80";
                txtInput18.Text = "80";
                txtInput19.Text = "2.60";
                txtInput20.Text = "30";
                txtInput21.Text = "0.08";
                txtInput22.Text = "75";
                txtOutput3.Text = "0";
                txtOutput7.Text = "25";
                txtOutput31.Text = "12";
                txtOutput27.Text = "0.7";
                txtOutput28.Text = "0.049";
                foreach (DataGridViewRow row1 in dataGridView1.Rows)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        row1.Cells[i].Value = string.Empty;
                    }
                }
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        row1.Cells[i].Value = string.Empty;
                    }
                }
                foreach (Control c in this.groupBox3.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox4.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox5.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox6.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
                foreach (Control c in this.groupBox9.Controls)
                {
                    if (c is TextBox)
                    {
                        //清掉含有TexBox控件上的内容
                        c.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        #endregion

        #region 导出
        private void button18_Click_1(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--分布式能源1.xlsx";
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

            #region   输入导出
            string[] d = new string[27];
            d[0] = comboBox1.Text;
            d[1] = txtInput2.Text;
            d[2] = txtInput3.Text;
            d[3] = txtInput4.Text;
            d[4] = txtInput5.Text;
            d[5] = txtInput6.Text;
            d[6] = txtInput7.Text;
            d[7] = txtInput8.Text;
            d[8] = txtInput9.Text;
            d[9] = txtInput10.Text;
            d[10] = txtInput11.Text;
            d[11] = txtOutput3.Text;
            d[12] = txtOutput7.Text;

            d[13] = txtInput12.Text;
            d[14] = txtInput13.Text;
            d[15] = txtInput14.Text;
            d[16] = txtInput15.Text;
            d[17] = txtInput16.Text;
            d[18] = txtInput17.Text;
            d[19] = txtInput18.Text;
            d[20] = txtInput19.Text;
            d[21] = txtInput20.Text;
            d[22] = txtInput21.Text;
            d[23] = txtInput22.Text;
            d[24] = txtOutput31.Text;
            d[25] = txtOutput27.Text;
            d[26] = txtOutput28.Text;
            for (int i = 0; i < 13; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(2).SetCellValue(d[i]);
            }
            for (int i = 0; i < 14; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(7).SetCellValue(d[13 + i]);
            }
            #endregion
            #region   投资/效益导出
            string[] c = new string[8];
            c[0] = txtOutput1.Text;
            c[1] = txtOutput2.Text;
            c[2] = txtOutput4.Text;

            c[3] = txtOutput5.Text;
            c[4] = txtOutput6.Text;
            c[5] = txtOutput8.Text;
            c[6] = txtOutput9.Text;
            c[7] = txtOutput10.Text;

            sheet1[0].GetRow(20).GetCell(2).SetCellValue(c[0]);
            sheet1[0].GetRow(21).GetCell(2).SetCellValue(c[1]);
            sheet1[0].GetRow(22).GetCell(2).SetCellValue(c[2]);
            for (int i = 0; i < 5; i++)
            {
                sheet1[0].GetRow(20 + i).GetCell(7).SetCellValue(c[3 + i]);
            }
            #endregion
            #region 收入/成本导出
            string[] b = new string[21];
            b[0] = txtOutput11.Text;
            b[1] = txtOutput12.Text;
            b[2] = txtOutput13.Text;
            b[3] = txtOutput14.Text;
            b[4] = txtOutput15.Text;
            b[5] = txtOutput16.Text;
            b[6] = txtOutput17.Text;
            b[7] = txtOutput18.Text;
            b[8] = txtOutput19.Text;
            b[9] = txtOutput20.Text;

            b[10] = txtOutput21.Text;
            b[11] = txtOutput22.Text;
            b[12] = txtOutput23.Text;
            b[13] = txtOutput24.Text;
            b[14] = txtOutput25.Text;
            b[15] = txtOutput26.Text;
            b[16] = txtOutput29.Text;
            b[17] = txtOutput30.Text;
            b[18] = txtOutput32.Text;
            b[19] = txtOutput33.Text;
            b[20] = txtOutput34.Text;

            for (int i = 0; i < 9; i++)
            {
                sheet1[0].GetRow(28 + i).GetCell(2).SetCellValue(b[i]);
            }
            for (int i = 10; i < 20; i++)
            {
                sheet1[0].GetRow(18 + i).GetCell(7).SetCellValue(b[i]);
            }
            #endregion
            #region  把datagridview中的数据导出到Excel
            string[] a = new string[32];
            for (int i = 0; i < 9; i++)
            {
                for (int j = 0; j < 32; j++)
                {
                    a[j] = (dataGridView1.Rows[j].Cells[i].Value).ToString();
                    sheet1[0].GetRow(43 + i).GetCell(4 + j).SetCellValue(a[j]);
                }
            }
            #endregion

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "分布式能源投资决策模型--技术经济型模型.cs";//设置默认文件名
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
        #endregion

        #region  1-边界电价
        private void button8_Click_1(object sender, EventArgs e)
        {

            try
            {
                #region  清除其他按钮

                this.textBox10.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox6.Text ="";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);
                #endregion
                #region  计算部分

                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
                double Var1 = 0.72;
                txtInput16.Text = Var1.ToString();                                   //单位电价（含税）
                textBox1.Text = Var1.ToString();
                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);
                Double M2 = Convert.ToDouble(txtInput11.Text);
                Double M3 = Convert.ToDouble(txtInput7.Text);
                Double M4 = Convert.ToDouble(txtInput16.Text);
                Double M5 = Convert.ToDouble(txtInput8.Text);
                Double M6 = Convert.ToDouble(txtInput17.Text);
                Double M7 = Convert.ToDouble(txtInput18.Text);
                Double M19 = Convert.ToDouble(txtInput9.Text);
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);
                Double M9 = Convert.ToDouble(txtInput10.Text);
                Double M10 = Convert.ToDouble(txtInput19.Text);
                Double M11 = Convert.ToDouble(txtInput20.Text);
                Double M12 = Convert.ToDouble(txtInput12.Text);
                Double M13 = Convert.ToDouble(txtOutput27.Text);
                Double M14 = Convert.ToDouble(txtOutput28.Text);
                Double M15 = Convert.ToDouble(txtOutput3.Text);
                Double M16 = Convert.ToDouble(txtInput22.Text);
                Double M17 = Convert.ToDouble(txtOutput31.Text);
                Double M18 = Convert.ToDouble(txtInput21.Text);
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税

                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出

                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");
                }  //列
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                CostCalculate();
                InvestCalculate();
                #endregion
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region    2-气价盈亏平衡点    

        private void button10_Click_1(object sender, EventArgs e)
        {

            try
            {
                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox10.Text = "";
                this.textBox6.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion
                #region   计算部分
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);
                Double M2 = Convert.ToDouble(txtInput11.Text);
                Double M3 = Convert.ToDouble(txtInput7.Text);
                Double M4 = Convert.ToDouble(txtInput16.Text);
                Double M5 = Convert.ToDouble(txtInput8.Text);
                Double M6 = Convert.ToDouble(txtInput17.Text);
                Double M7 = Convert.ToDouble(txtInput18.Text);
                Double M19 = Convert.ToDouble(txtInput9.Text);
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);
                Double M9 = Convert.ToDouble(txtInput10.Text);
                double Var2 = 3.69;
                textBox7.Text = Var2.ToString();
                txtInput19.Text = Var2.ToString();
                Double M10 = Convert.ToDouble(txtInput19.Text);
                Double M11 = Convert.ToDouble(txtInput20.Text);
                Double M12 = Convert.ToDouble(txtInput12.Text);
                Double M13 = Convert.ToDouble(txtOutput27.Text);
                Double M14 = Convert.ToDouble(txtOutput28.Text);
                Double M15 = Convert.ToDouble(txtOutput3.Text);
                Double M16 = Convert.ToDouble(txtInput22.Text);
                Double M17 = Convert.ToDouble(txtOutput31.Text);
                Double M18 = Convert.ToDouble(txtInput21.Text);
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
                //数据表中第一行的内容计算
                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 

                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;

                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                //数据表中第二行及其以后的数据计算
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税

                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出


                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                }
                #endregion
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                CostCalculate();
                    InvestCalculate();
                    #endregion
                
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }




        #endregion

        #region   2-边界气价
        private void button14_Click_1(object sender, EventArgs e)
        {
            try
            {

                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox6.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox10.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion
                #region  计算部分
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);
                Double M2 = Convert.ToDouble(txtInput11.Text);
                Double M3 = Convert.ToDouble(txtInput7.Text);
                Double M4 = Convert.ToDouble(txtInput16.Text);
                Double M5 = Convert.ToDouble(txtInput8.Text);
                Double M6 = Convert.ToDouble(txtInput17.Text);
                Double M7 = Convert.ToDouble(txtInput18.Text);
                Double M19 = Convert.ToDouble(txtInput9.Text);
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);
                Double M9 = Convert.ToDouble(txtInput10.Text);
                double Var2 = 2.80;
                textBox2.Text = Var2.ToString();
                txtInput19.Text = Var2.ToString();
                Double M10 = Convert.ToDouble(txtInput19.Text);
                Double M11 = Convert.ToDouble(txtInput20.Text);
                Double M12 = Convert.ToDouble(txtInput12.Text);
                Double M13 = Convert.ToDouble(txtOutput27.Text);
                Double M14 = Convert.ToDouble(txtOutput28.Text);
                Double M15 = Convert.ToDouble(txtOutput3.Text);
                Double M16 = Convert.ToDouble(txtInput22.Text);
                Double M17 = Convert.ToDouble(txtOutput31.Text);
                Double M18 = Convert.ToDouble(txtInput21.Text);
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
                //数据表中第一行的内容计算
                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                //数据表中第二行及其以后的数据计算
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税

                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出

                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #endregion
                }
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                CostCalculate();
                InvestCalculate();
                #endregion
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region  4-年利用小时数盈亏平衡点
        private void button12_Click_1(object sender, EventArgs e)
        {
            try
            {
                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox10.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox6.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion
                #region   计算部分
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
                //double Var1 = 0.57;
                //txtInput16.Text = Var1.ToString();                                   //单位电价（含税）
                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M2 = Convert.ToDouble(txtInput11.Text); //负荷率：
                Double M3 = Convert.ToDouble(txtInput7.Text);// 发电机组装机容量：
                Double M4 = Convert.ToDouble(txtInput16.Text);//单位电价：
                Double M5 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量：
                Double M6 = Convert.ToDouble(txtInput17.Text);//单位热价：
                Double M7 = Convert.ToDouble(txtInput18.Text);//单位冷价：
                Double M19 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量：
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);
                double Var1 = 1823;
                txtInput10.Text = Var1.ToString();
                textBox9.Text = Var1.ToString();
                Double M9 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M10 = Convert.ToDouble(txtInput19.Text);//单位气价：
                Double M11 = Convert.ToDouble(txtInput20.Text);//折旧年限：
                Double M12 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资：
                Double M13 = Convert.ToDouble(txtOutput27.Text);//贷款比例：
                Double M14 = Convert.ToDouble(txtOutput28.Text); //贷款利率：
                Double M15 = Convert.ToDouble(txtOutput3.Text);//单位补贴：
                Double M16 = Convert.ToDouble(txtInput22.Text);//站场定员：
                Double M17 = Convert.ToDouble(txtOutput31.Text);// 人员费用：
                Double M18 = Convert.ToDouble(txtInput21.Text);// 单位运维成本：
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本

                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #endregion

                }
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                CostCalculate();
                InvestCalculate();
                #endregion

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region    4-边界利用小时数
        private void button11_Click_1(object sender, EventArgs e)
        {
            try
            {
                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox6.Text = "";
                this.textBox10.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion
                #region  计算部分
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M2 = Convert.ToDouble(txtInput11.Text); //负荷率：
                Double M3 = Convert.ToDouble(txtInput7.Text);// 发电机组装机容量：
                Double M4 = Convert.ToDouble(txtInput16.Text);//单位电价：
                Double M5 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量：
                Double M6 = Convert.ToDouble(txtInput17.Text);//单位热价：
                Double M7 = Convert.ToDouble(txtInput18.Text);//单位冷价：
                Double M19 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量：
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);
                double Var1 = 4828;
                txtInput10.Text = Var1.ToString();
                textBox5.Text = Var1.ToString();
                Double M9 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M10 = Convert.ToDouble(txtInput19.Text);//单位气价：
                Double M11 = Convert.ToDouble(txtInput20.Text);//折旧年限：
                Double M12 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资：
                Double M13 = Convert.ToDouble(txtOutput27.Text);//贷款比例：
                Double M14 = Convert.ToDouble(txtOutput28.Text); //贷款利率：
                Double M15 = Convert.ToDouble(txtOutput3.Text);//单位补贴：
                Double M16 = Convert.ToDouble(txtInput22.Text);//站场定员：
                Double M17 = Convert.ToDouble(txtOutput31.Text);// 人员费用：
                Double M18 = Convert.ToDouble(txtInput21.Text);// 单位运维成本：
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本

                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出

                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #endregion
                }
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                CostCalculate();
                InvestCalculate();
                #endregion
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        #region   5-单位投资盈亏平衡
        private void button13_Click_1(object sender, EventArgs e)
        {
            try
            {
                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox6.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion
                #region   ji算部分
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                double Var1 = 20234;
                txtInput12.Text = Var1.ToString();
                textBox10.Text = Var1.ToString();
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资

                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M2 = Convert.ToDouble(txtInput11.Text); //负荷率：
                Double M3 = Convert.ToDouble(txtInput7.Text);// 发电机组装机容量：
                Double M4 = Convert.ToDouble(txtInput16.Text);//单位电价：
                Double M5 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量：
                Double M6 = Convert.ToDouble(txtInput17.Text);//单位热价：
                Double M7 = Convert.ToDouble(txtInput18.Text);//单位冷价：
                Double M19 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量：
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);

                Double M9 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M10 = Convert.ToDouble(txtInput19.Text);//单位气价：
                Double M11 = Convert.ToDouble(txtInput20.Text);//折旧年限：
                Double M12 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资：
                Double M13 = Convert.ToDouble(txtOutput27.Text);//贷款比例：
                Double M14 = Convert.ToDouble(txtOutput28.Text); //贷款利率：
                Double M15 = Convert.ToDouble(txtOutput3.Text);//单位补贴：
                Double M16 = Convert.ToDouble(txtInput22.Text);//站场定员：
                Double M17 = Convert.ToDouble(txtOutput31.Text);// 人员费用：
                Double M18 = Convert.ToDouble(txtInput21.Text);// 单位运维成本：
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本

                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出

                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #endregion
                }
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                CostCalculate();
                InvestCalculate();
                #endregion
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region  5-边界单位投资
        private void button7_Click_1(object sender, EventArgs e)
        {
            try
            {
                #region  清除其他按钮

                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox6.Text = "";
                this.textBox5.Text = "";
                this.textBox10.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";

                #endregion
                #region   错误检测部分
                string str1 = label7.Text;
                string str2 = txtInput7.Text; //发电机组装机容量：
                string str3 = label8.Text;
                string str4 = txtInput8.Text; //制热机组装机容量：
                string str5 = label9.Text;
                string str6 = txtInput9.Text; //制冷机组装机容量：
                string str7 = label10.Text;
                string str8 = txtInput10.Text;//年利用小时数：
                string str9 = label11.Text;
                string str10 = txtInput11.Text;//负荷率：
                string str11 = label12.Text;
                string str12 = txtInput12.Text;//单位电力装机投资：
                string str13 = label13.Text;
                string str14 = txtInput13.Text; //单位装机补贴：
                string str15 = label20.Text;
                string str16 = txtInput14.Text;//发电气耗
                string str17 = label23.Text;
                string str18 = txtInput15.Text;//供热气耗
                string str19 = label21.Text;
                string str20 = txtInput16.Text;//单位电价
                string str21 = label29.Text; //单位热价
                string str22 = txtInput17.Text;
                string str23 = label28.Text;
                string str24 = txtInput18.Text;//单位冷价
                string str25 = label27.Text;
                string str26 = txtInput19.Text;//单位气价
                string str27 = label26.Text;
                string str28 = txtInput20.Text;//折旧年限
                string str29 = label25.Text;
                string str30 = txtInput21.Text;//单位运维成本
                string str31 = label22.Text;
                string str32 = txtInput22.Text;//站场定员
                string str33 = label104.Text;
                string str34 = txtOutput31.Text;//人员费用：
                string str35 = label74.Text;
                string str36 = txtOutput27.Text;//贷款比例
                string str37 = label73.Text;
                string str38 = txtOutput28.Text;//贷款利率
                string str39 = label44.Text;
                string str40 = txtOutput3.Text;//：单位补贴
                Common.ParameterErrorDetectionDynamoValume(str1, str2);
                Common.ParameterErrorDetectionHeatProductValume(str3, str4);
                Common.ParameterErrorDetectionColdProductValume(str5, str6);
                Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
                Common.ParameterErrorDetectionBurthenRatio(str9, str10);
                Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
                Common.ParameterErrorDetectionSingleAllowance(str13, str14);
                Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
                Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
                Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
                Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
                Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
                Common.ParameterErrorDetection(str25, str26);
                Common.ParameterErrorDetection(str27, str28);
                Common.ParameterErrorDetection(str29, str30);
                Common.ParameterErrorDetection(str31, str32);
                Common.ParameterErrorDetection(str33, str34);
                Common.ParameterErrorDetection(str35, str36);
                Common.ParameterErrorDetection(str37, str38);
                Common.ParameterErrorDetectionColdProductValume(str39, str40);

                #endregion
                #region   计算部分
                Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
                double Var1 = 6546;
                txtInput12.Text = Var1.ToString();
                textBox4.Text = Var1.ToString();
                Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资

                double Var113 = P1 * P2 / 10000;                                     //1-总投资  

                Double M1 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M2 = Convert.ToDouble(txtInput11.Text); //负荷率：
                Double M3 = Convert.ToDouble(txtInput7.Text);// 发电机组装机容量：
                Double M4 = Convert.ToDouble(txtInput16.Text);//单位电价：
                Double M5 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量：
                Double M6 = Convert.ToDouble(txtInput17.Text);//单位热价：
                Double M7 = Convert.ToDouble(txtInput18.Text);//单位冷价：
                Double M19 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量：
                Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
                Double M8 = Convert.ToDouble(txtInput14.Text);

                Double M9 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
                Double M10 = Convert.ToDouble(txtInput19.Text);//单位气价：
                Double M11 = Convert.ToDouble(txtInput20.Text);//折旧年限：
                Double M12 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资：
                Double M13 = Convert.ToDouble(txtOutput27.Text);//贷款比例：
                Double M14 = Convert.ToDouble(txtOutput28.Text); //贷款利率：
                Double M15 = Convert.ToDouble(txtOutput3.Text);//单位补贴：
                Double M16 = Convert.ToDouble(txtInput22.Text);//站场定员：
                Double M17 = Convert.ToDouble(txtOutput31.Text);// 人员费用：
                Double M18 = Convert.ToDouble(txtInput21.Text);// 单位运维成本：
                Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本

                Double Var123 = Var108 - Var122;  // 2-所得税
                this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
                this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
                this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税
                //空白填充为零
                this.dataGridView1.Rows[0].Cells[0].Value = 0;
                this.dataGridView1.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[3].Value = 0;
                    this.dataGridView1.Rows[i].Cells[7].Value = 0;
                    this.dataGridView1.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView1.Rows[31].Cells[8].Value = 0;
                #region 二次填充
                //空白填充为零
                this.dataGridView2.Rows[0].Cells[0].Value = 0;
                this.dataGridView2.Rows[0].Cells[1].Value = 0;
                for (int i = 0; i < 5; i++)
                {
                    this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
                }
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView2.Rows[i].Cells[3].Value = 0;
                    this.dataGridView2.Rows[i].Cells[7].Value = 0;
                    this.dataGridView2.Rows[i].Cells[8].Value = 0;
                }
                this.dataGridView2.Rows[31].Cells[8].Value = 0;
                #endregion
                for (int i = 1; i < 31; i++)
                {
                    this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                    Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                    this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    this.dataGridView1.Rows[0].Cells[0].Value = 0;
                    this.dataGridView1.Rows[0].Cells[1].Value = 0;
                    #region 二次写入
                    this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                    this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                    this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                    this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                    this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                    Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                    Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                    this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                    #endregion

                }
                #region  合计
                string[] a = new string[16];
                for (int i = 0; i < 16; i++)
                {
                    a[i] = Convert.ToString(0);
                }
                for (int m = 0; m < 31; m++)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                    for (int i = 0; i < 8; i++)
                    {
                        a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                    this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
                }
                #endregion
                CostCalculate();
                InvestCalculate();
                #endregion   计算
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region  3-热价盈亏平衡点
        private void button6_Click(object sender, EventArgs e)
        {
            #region  清除其他按钮

            this.textBox1.Text = "";
            this.textBox2.Text = "";
            this.textBox3.Text = "";
            this.textBox4.Text = "";
            this.textBox5.Text = "";
            this.textBox10.Text = "";
            this.textBox7.Text = "";
            this.textBox6.Text = "";
            this.textBox9.Text = "";

            #endregion
            #region   错误检测部分
            string str1 = label7.Text;
            string str2 = txtInput7.Text; //发电机组装机容量：
            string str3 = label8.Text;
            string str4 = txtInput8.Text; //制热机组装机容量：
            string str5 = label9.Text;
            string str6 = txtInput9.Text; //制冷机组装机容量：
            string str7 = label10.Text;
            string str8 = txtInput10.Text;//年利用小时数：
            string str9 = label11.Text;
            string str10 = txtInput11.Text;//负荷率：
            string str11 = label12.Text;
            string str12 = txtInput12.Text;//单位电力装机投资：
            string str13 = label13.Text;
            string str14 = txtInput13.Text; //单位装机补贴：
            string str15 = label20.Text;
            string str16 = txtInput14.Text;//发电气耗
            string str17 = label23.Text;
            string str18 = txtInput15.Text;//供热气耗
            string str19 = label21.Text;
            string str20 = txtInput16.Text;//单位电价
            string str21 = label29.Text; //单位热价
            string str22 = txtInput17.Text;
            string str23 = label28.Text;
            string str24 = txtInput18.Text;//单位冷价
            string str25 = label27.Text;
            string str26 = txtInput19.Text;//单位气价
            string str27 = label26.Text;
            string str28 = txtInput20.Text;//折旧年限
            string str29 = label25.Text;
            string str30 = txtInput21.Text;//单位运维成本
            string str31 = label22.Text;
            string str32 = txtInput22.Text;//站场定员
            string str33 = label104.Text;
            string str34 = txtOutput31.Text;//人员费用：
            string str35 = label74.Text;
            string str36 = txtOutput27.Text;//贷款比例
            string str37 = label73.Text;
            string str38 = txtOutput28.Text;//贷款利率
            string str39 = label44.Text;
            string str40 = txtOutput3.Text;//：单位补贴
            Common.ParameterErrorDetectionDynamoValume(str1, str2);
            Common.ParameterErrorDetectionHeatProductValume(str3, str4);
            Common.ParameterErrorDetectionColdProductValume(str5, str6);
            Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
            Common.ParameterErrorDetectionBurthenRatio(str9, str10);
            Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
            Common.ParameterErrorDetectionSingleAllowance(str13, str14);
            Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
            Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
            Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
            Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
            Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
            Common.ParameterErrorDetection(str25, str26);
            Common.ParameterErrorDetection(str27, str28);
            Common.ParameterErrorDetection(str29, str30);
            Common.ParameterErrorDetection(str31, str32);
            Common.ParameterErrorDetection(str33, str34);
            Common.ParameterErrorDetection(str35, str36);
            Common.ParameterErrorDetection(str37, str38);
            Common.ParameterErrorDetectionColdProductValume(str39, str40);
            #endregion
            #region  计算部分
            double Var1 = -538;
            txtInput17.Text = Var1.ToString();                                   //单位电价（含税）
            textBox8.Text = Var1.ToString();
            Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
            Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资

            double Var113 = P1 * P2 / 10000;                                     //1-总投资  

            Double M1 = Convert.ToDouble(txtInput10.Text);
            Double M2 = Convert.ToDouble(txtInput11.Text);
            Double M3 = Convert.ToDouble(txtInput7.Text);
            Double M4 = Convert.ToDouble(txtInput16.Text);
            Double M5 = Convert.ToDouble(txtInput8.Text);
            Double M6 = Convert.ToDouble(txtInput17.Text);
            Double M7 = Convert.ToDouble(txtInput18.Text);
            Double M19 = Convert.ToDouble(txtInput9.Text);
            Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
            Double M8 = Convert.ToDouble(txtInput14.Text);
            Double M9 = Convert.ToDouble(txtInput10.Text);
            Double M10 = Convert.ToDouble(txtInput19.Text);
            Double M11 = Convert.ToDouble(txtInput20.Text);
            Double M12 = Convert.ToDouble(txtInput12.Text);
            Double M13 = Convert.ToDouble(txtOutput27.Text);
            Double M14 = Convert.ToDouble(txtOutput28.Text);
            Double M15 = Convert.ToDouble(txtOutput3.Text);
            Double M16 = Convert.ToDouble(txtInput22.Text);
            Double M17 = Convert.ToDouble(txtOutput31.Text);
            Double M18 = Convert.ToDouble(txtInput21.Text);
            Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
            Double Var123 = Var108 - Var122;  // 2-所得税
            this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
            this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
            this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税
                                                                               //空白填充为零
            this.dataGridView1.Rows[0].Cells[0].Value = 0;
            this.dataGridView1.Rows[0].Cells[1].Value = 0;
            for (int i = 0; i < 5; i++)
            {
                this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
            }
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView1.Rows[i].Cells[3].Value = 0;
                this.dataGridView1.Rows[i].Cells[7].Value = 0;
                this.dataGridView1.Rows[i].Cells[8].Value = 0;
            }
            this.dataGridView1.Rows[31].Cells[8].Value = 0;
            #region 二次填充
            //空白填充为零
            this.dataGridView2.Rows[0].Cells[0].Value = 0;
            this.dataGridView2.Rows[0].Cells[1].Value = 0;
            for (int i = 0; i < 5; i++)
            {
                this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
            }
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView2.Rows[i].Cells[3].Value = 0;
                this.dataGridView2.Rows[i].Cells[7].Value = 0;
                this.dataGridView2.Rows[i].Cells[8].Value = 0;
            }
            this.dataGridView2.Rows[31].Cells[8].Value = 0;
            #endregion
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税

                this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出
                #region 二次写入
                this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                #endregion
            }  //列
            #region  合计
            string[] a = new string[16];
            for (int i = 0; i < 16; i++)
            {
                a[i] = Convert.ToString(0);
            }
            for (int m = 0; m < 31; m++)
            {
                for (int i = 0; i < 8; i++)
                {
                    a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                }
                for (int i = 0; i < 8; i++)
                {
                    a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                }
            }
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
            }
            #endregion
            CostCalculate();
            InvestCalculate();
            #endregion
        }
        #endregion

        #region  3-边界热价
        private void button15_Click(object sender, EventArgs e)
        {
            #region  清除其他按钮

            this.textBox1.Text = "";
            this.textBox2.Text = "";
            this.textBox6.Text = "";
            this.textBox4.Text = "";
            this.textBox5.Text = "";
            this.textBox10.Text = "";
            this.textBox7.Text = "";
            this.textBox8.Text = "";
            this.textBox9.Text = "";

            #endregion
            #region   错误检测部分
            string str1 = label7.Text;
            string str2 = txtInput7.Text; //发电机组装机容量：
            string str3 = label8.Text;
            string str4 = txtInput8.Text; //制热机组装机容量：
            string str5 = label9.Text;
            string str6 = txtInput9.Text; //制冷机组装机容量：
            string str7 = label10.Text;
            string str8 = txtInput10.Text;//年利用小时数：
            string str9 = label11.Text;
            string str10 = txtInput11.Text;//负荷率：
            string str11 = label12.Text;
            string str12 = txtInput12.Text;//单位电力装机投资：
            string str13 = label13.Text;
            string str14 = txtInput13.Text; //单位装机补贴：
            string str15 = label20.Text;
            string str16 = txtInput14.Text;//发电气耗
            string str17 = label23.Text;
            string str18 = txtInput15.Text;//供热气耗
            string str19 = label21.Text;
            string str20 = txtInput16.Text;//单位电价
            string str21 = label29.Text; //单位热价
            string str22 = txtInput17.Text;
            string str23 = label28.Text;
            string str24 = txtInput18.Text;//单位冷价
            string str25 = label27.Text;
            string str26 = txtInput19.Text;//单位气价
            string str27 = label26.Text;
            string str28 = txtInput20.Text;//折旧年限
            string str29 = label25.Text;
            string str30 = txtInput21.Text;//单位运维成本
            string str31 = label22.Text;
            string str32 = txtInput22.Text;//站场定员
            string str33 = label104.Text;
            string str34 = txtOutput31.Text;//人员费用：
            string str35 = label74.Text;
            string str36 = txtOutput27.Text;//贷款比例
            string str37 = label73.Text;
            string str38 = txtOutput28.Text;//贷款利率
            string str39 = label44.Text;
            string str40 = txtOutput3.Text;//：单位补贴
            Common.ParameterErrorDetectionDynamoValume(str1, str2);
            Common.ParameterErrorDetectionHeatProductValume(str3, str4);
            Common.ParameterErrorDetectionColdProductValume(str5, str6);
            Common.ParameterErrorDetectionAnnualUsedHour(str7, str8);
            Common.ParameterErrorDetectionBurthenRatio(str9, str10);
            Common.ParameterErrorDetectionSingleElectrityInvest(str11, str12);
            Common.ParameterErrorDetectionSingleAllowance(str13, str14);
            Common.ParameterErrorDetectionElectrityProductGasUsed(str15, str16);
            Common.ParameterErrorDetectionHeatingProductGasUsed(str17, str18);
            Common.ParameterErrorDetectionSingleElectrityPrice(str19, str20);
            Common.ParameterErrorDetectionPerHeaatingPrice(str21, str22);
            Common.ParameterErrorDetectionPerHeaatingPrice(str23, str24);
            Common.ParameterErrorDetection(str25, str26);
            Common.ParameterErrorDetection(str27, str28);
            Common.ParameterErrorDetection(str29, str30);
            Common.ParameterErrorDetection(str31, str32);
            Common.ParameterErrorDetection(str33, str34);
            Common.ParameterErrorDetection(str35, str36);
            Common.ParameterErrorDetection(str37, str38);
            Common.ParameterErrorDetectionColdProductValume(str39, str40);
            #endregion
            #region  计算部分

            Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
            Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
            double Var1 = -33;
            txtInput17.Text = Var1.ToString();                                   //单位电价（含税）
            textBox3.Text = Var1.ToString();
            double Var113 = P1 * P2 / 10000;                                     //1-总投资  

            Double M1 = Convert.ToDouble(txtInput10.Text);
            Double M2 = Convert.ToDouble(txtInput11.Text);
            Double M3 = Convert.ToDouble(txtInput7.Text);
            Double M4 = Convert.ToDouble(txtInput16.Text);
            Double M5 = Convert.ToDouble(txtInput8.Text);
            Double M6 = Convert.ToDouble(txtInput17.Text);
            Double M7 = Convert.ToDouble(txtInput18.Text);
            Double M19 = Convert.ToDouble(txtInput9.Text);
            Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
            Double M8 = Convert.ToDouble(txtInput14.Text);
            Double M9 = Convert.ToDouble(txtInput10.Text);
            Double M10 = Convert.ToDouble(txtInput19.Text);
            Double M11 = Convert.ToDouble(txtInput20.Text);
            Double M12 = Convert.ToDouble(txtInput12.Text);
            Double M13 = Convert.ToDouble(txtOutput27.Text);
            Double M14 = Convert.ToDouble(txtOutput28.Text);
            Double M15 = Convert.ToDouble(txtOutput3.Text);
            Double M16 = Convert.ToDouble(txtInput22.Text);
            Double M17 = Convert.ToDouble(txtOutput31.Text);
            Double M18 = Convert.ToDouble(txtInput21.Text);
            Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
            Double Var123 = Var108 - Var122;  // 2-所得税
            this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
            this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
            this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税
                                                                               //空白填充为零
            this.dataGridView1.Rows[0].Cells[0].Value = 0;
            this.dataGridView1.Rows[0].Cells[1].Value = 0;
            for (int i = 0; i < 5; i++)
            {
                this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
            }
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView1.Rows[i].Cells[3].Value = 0;
                this.dataGridView1.Rows[i].Cells[7].Value = 0;
                this.dataGridView1.Rows[i].Cells[8].Value = 0;
            }
            this.dataGridView1.Rows[31].Cells[8].Value = 0;
            #region 二次填充
            //空白填充为零
            this.dataGridView2.Rows[0].Cells[0].Value = 0;
            this.dataGridView2.Rows[0].Cells[1].Value = 0;
            for (int i = 0; i < 5; i++)
            {
                this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
            }
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView2.Rows[i].Cells[3].Value = 0;
                this.dataGridView2.Rows[i].Cells[7].Value = 0;
                this.dataGridView2.Rows[i].Cells[8].Value = 0;
            }
            this.dataGridView2.Rows[31].Cells[8].Value = 0;
            #endregion
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税

                this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出
                #region 二次写入
                this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                #endregion
            }  //列
            #region  合计
            string[] a = new string[16];
            for (int i = 0; i < 16; i++)
            {
                a[i] = Convert.ToString(0);
            }
            for (int m = 0; m < 31; m++)
            {
                for (int i = 0; i < 8; i++)
                {
                    a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                }
                for (int i = 0; i < 8; i++)
                {
                    a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                }
            }
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
            }
            #endregion
            CostCalculate();
            InvestCalculate();
            #endregion

        }
        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            try { 
            CostCalculate();
            InvestCalculate();

            Double P1 = Convert.ToDouble(txtInput7.Text);                //发电机组装机容量
            Double P2 = Convert.ToDouble(txtInput12.Text);                //单位电力装机投资
            double Var113 = P1 * P2 / 10000;                                     //1-总投资  

            Double M1 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
            Double M2 = Convert.ToDouble(txtInput11.Text); //负荷率：
            Double M3 = Convert.ToDouble(txtInput7.Text);// 发电机组装机容量：
            Double M4 = Convert.ToDouble(txtInput16.Text);//单位电价：
            Double M5 = Convert.ToDouble(txtInput8.Text);//制热机组装机容量：
            Double M6 = Convert.ToDouble(txtInput17.Text);//单位热价：
            Double M7 = Convert.ToDouble(txtInput18.Text);//单位冷价：
            Double M19 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量：
            Double Var108 = Calculatemath.AllSum(M1, M2, M3, M4, M5, M6, M7, M19);//总收入
            Double M8 = Convert.ToDouble(txtInput14.Text);
            Double M9 = Convert.ToDouble(txtInput10.Text);//年利用小时数：
            Double M10 = Convert.ToDouble(txtInput19.Text);//单位气价：
            Double M11 = Convert.ToDouble(txtInput20.Text);//折旧年限：
            Double M12 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资：
            Double M13 = Convert.ToDouble(txtOutput27.Text);//贷款比例：
            Double M14 = Convert.ToDouble(txtOutput28.Text); //贷款利率：
            Double M15 = Convert.ToDouble(txtOutput3.Text);//单位补贴：
            Double M16 = Convert.ToDouble(txtInput22.Text);//站场定员：
            Double M17 = Convert.ToDouble(txtOutput31.Text);// 人员费用：
            Double M18 = Convert.ToDouble(txtInput21.Text);// 单位运维成本：
            Double Var122 = Calculatemath.AllCost(M8, M2, M9, M3, M10, M11, M12, M13, M14, M15, M16, M17, M18);//总成本
            InvestCalculate();
            CostCalculate();
            Double Var123 = Var108 - Var122;  // 2-所得税
            this.dataGridView1.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
            this.dataGridView1.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
            this.dataGridView1.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 

            this.dataGridView2.Rows[0].Cells[0].Value = this.dataGridView1.Rows[0].Cells[1].Value;//1行-现金流入=1行-营业收入
            this.dataGridView2.Rows[0].Cells[3].Value = Var113.ToString("0");//1行-建设投资
            this.dataGridView2.Rows[0].Cells[2].Value = (Var113).ToString("0");//现金流出=1行-建设投资+1行-成本费用+1行-所得税 

            #region 填充为零
            //空白填充为零
            this.dataGridView1.Rows[0].Cells[0].Value = 0;
            this.dataGridView1.Rows[0].Cells[1].Value = 0;
            for (int i = 0; i < 5; i++)
            {
                this.dataGridView1.Rows[0].Cells[4 + i].Value = 0;
            }
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView1.Rows[i].Cells[3].Value = 0;
                this.dataGridView1.Rows[i].Cells[7].Value = 0;
                this.dataGridView1.Rows[i].Cells[8].Value = 0;
            }
            this.dataGridView1.Rows[31].Cells[8].Value = 0;
            #region 二次填充
            //空白填充为零
            this.dataGridView2.Rows[0].Cells[0].Value = 0;
            this.dataGridView2.Rows[0].Cells[1].Value = 0;
            for (int i = 0; i < 5; i++)
            {
                this.dataGridView2.Rows[0].Cells[4 + i].Value = 0;
            }
            for (int i = 1; i < 31; i++)
            {
                this.dataGridView2.Rows[i].Cells[3].Value = 0;
                this.dataGridView2.Rows[i].Cells[7].Value = 0;
                this.dataGridView2.Rows[i].Cells[8].Value = 0;
            }
            this.dataGridView2.Rows[31].Cells[8].Value = 0;
            #endregion
            #endregion

            for (int i = 1; i < 31; i++)
            {
                this.dataGridView1.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                this.dataGridView1.Rows[i].Cells[0].Value = this.dataGridView1.Rows[i].Cells[1].Value;//现金流入=营业收入
                this.dataGridView1.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                this.dataGridView1.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                this.dataGridView1.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                Double B1 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[0].Value);
                Double B2 = Convert.ToDouble(this.dataGridView1.Rows[i].Cells[2].Value);
                this.dataGridView1.Rows[i].Cells[6].Value = (B1 - B2).ToString("0");                     //税后净现金流量=现金流入-现金流出
                #region 二次写入
                this.dataGridView2.Rows[i].Cells[1].Value = Var108.ToString("0"); //营业收入
                this.dataGridView2.Rows[i].Cells[0].Value = this.dataGridView2.Rows[i].Cells[1].Value;//现金流入=营业收入
                this.dataGridView2.Rows[i].Cells[2].Value = Var122.ToString("0");                    //现金流出=建设投资+成本费用+所得税
                this.dataGridView2.Rows[i].Cells[4].Value = Var122.ToString("0");                     //成本费用
                this.dataGridView2.Rows[i].Cells[5].Value = Var123.ToString("0");                     //所得税
                Double B3 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[0].Value);
                Double B4 = Convert.ToDouble(this.dataGridView2.Rows[i].Cells[2].Value);
                this.dataGridView2.Rows[i].Cells[6].Value = (B3 - B4).ToString("0");                     //税后净现金流量=现金流入-现金流出
                #endregion
            }
            #region  合计
            string[] a = new string[16];
            for (int i = 0; i < 16; i++)
            {
                a[i] = Convert.ToString(0);
            }
            for (int m = 0; m < 31; m++)
            {
                for (int i = 0; i < 8; i++)
                {
                    a[i] = (Convert.ToDouble(a[i]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                }
                for (int i = 0; i < 8; i++)
                {
                    a[8 + i] = (Convert.ToDouble(a[i + 8]) + Convert.ToDouble(this.dataGridView1.Rows[m].Cells[i].Value)).ToString();
                }
            }
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView1.Rows[31].Cells[i].Value = a[i].ToString();
                this.dataGridView2.Rows[31].Cells[i].Value = a[8 + i].ToString();
            }
            #endregion

        }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }






}

        private void button4_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.E)
            {
                button4.PerformClick();
            }
        }

        private void button9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button9.PerformClick();
            }
        }

        private void button16_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.E)
            {
                button16.PerformClick();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--分布式能源1.xlsx";
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

            #region   输入导出
            string[] d = new string[27];
            d[0] = comboBox1.Text;
            d[1] = txtInput2.Text;
            d[2] = txtInput3.Text;
            d[3] = txtInput4.Text;
            d[4] = txtInput5.Text;
            d[5] = txtInput6.Text;
            d[6] = txtInput7.Text;
            d[7] = txtInput8.Text;
            d[8] = txtInput9.Text;
            d[9] = txtInput10.Text;
            d[10] = txtInput11.Text;
            d[11] = txtOutput3.Text;
            d[12] = txtOutput7.Text;

            d[13] = txtInput12.Text;
            d[14] = txtInput13.Text;
            d[15] = txtInput14.Text;
            d[16] = txtInput15.Text;
            d[17] = txtInput16.Text;
            d[18] = txtInput17.Text;
            d[19] = txtInput18.Text;
            d[20] = txtInput19.Text;
            d[21] = txtInput20.Text;
            d[22] = txtInput21.Text;
            d[23] = txtInput22.Text;
            d[24] = txtOutput31.Text;
            d[25] = txtOutput27.Text;
            d[26] = txtOutput28.Text;
            for (int i = 0; i < 13; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(2).SetCellValue(d[i]);
            }
            for (int i = 0; i < 14; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(7).SetCellValue(d[13 + i]);
            }
            #endregion
            #region   投资/效益导出
            string[] c = new string[8];
            c[0] = txtOutput1.Text;
            c[1] = txtOutput2.Text;
            c[2] = txtOutput4.Text;

            c[3] = txtOutput5.Text;
            c[4] = txtOutput6.Text;
            c[5] = txtOutput8.Text;
            c[6] = txtOutput9.Text;
            c[7] = txtOutput10.Text;

            sheet1[0].GetRow(20).GetCell(2).SetCellValue(c[0]);
            sheet1[0].GetRow(21).GetCell(2).SetCellValue(c[1]);
            sheet1[0].GetRow(22).GetCell(2).SetCellValue(c[2]);
            for (int i = 0; i < 5; i++)
            {
                sheet1[0].GetRow(20 + i).GetCell(7).SetCellValue(c[3 + i]);
            }
            #endregion
            #region 收入/成本导出
            string[] b = new string[21];
            b[0] = txtOutput11.Text;
            b[1] = txtOutput12.Text;
            b[2] = txtOutput13.Text;
            b[3] = txtOutput14.Text;
            b[4] = txtOutput15.Text;
            b[5] = txtOutput16.Text;
            b[6] = txtOutput17.Text;
            b[7] = txtOutput18.Text;
            b[8] = txtOutput19.Text;
            b[9] = txtOutput20.Text;

            b[10] = txtOutput21.Text;
            b[11] = txtOutput22.Text;
            b[12] = txtOutput23.Text;
            b[13] = txtOutput24.Text;
            b[14] = txtOutput25.Text;
            b[15] = txtOutput26.Text;
            b[16] = txtOutput29.Text;
            b[17] = txtOutput30.Text;
            b[18] = txtOutput32.Text;
            b[19] = txtOutput33.Text;
            b[20] = txtOutput34.Text;

            for (int i = 0; i < 9; i++)
            {
                sheet1[0].GetRow(28 + i).GetCell(2).SetCellValue(b[i]);
            }
            for (int i = 10; i < 20; i++)
            {
                sheet1[0].GetRow(18 + i).GetCell(7).SetCellValue(b[i]);
            }
            #endregion
            #region  把datagridview中的数据导出到Excel
            string[] a = new string[32];
            for (int i = 0; i < 9; i++)
            {
                for (int j = 0; j < 32; j++)
                {
                    a[j] = (dataGridView1.Rows[j].Cells[i].Value).ToString();
                    sheet1[0].GetRow(43 + i).GetCell(4 + j).SetCellValue(a[j]);
                }
            }
            #endregion

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "分布式能源投资决策模型--技术经济型模型.cs";//设置默认文件名
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

        private void button9_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--分布式能源1.xlsx";
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

            #region   输入导出
            string[] d = new string[27];
            d[0] = comboBox1.Text;
            d[1] = txtInput2.Text;
            d[2] = txtInput3.Text;
            d[3] = txtInput4.Text;
            d[4] = txtInput5.Text;
            d[5] = txtInput6.Text;
            d[6] = txtInput7.Text;
            d[7] = txtInput8.Text;
            d[8] = txtInput9.Text;
            d[9] = txtInput10.Text;
            d[10] = txtInput11.Text;
            d[11] = txtOutput3.Text;
            d[12] = txtOutput7.Text;

            d[13] = txtInput12.Text;
            d[14] = txtInput13.Text;
            d[15] = txtInput14.Text;
            d[16] = txtInput15.Text;
            d[17] = txtInput16.Text;
            d[18] = txtInput17.Text;
            d[19] = txtInput18.Text;
            d[20] = txtInput19.Text;
            d[21] = txtInput20.Text;
            d[22] = txtInput21.Text;
            d[23] = txtInput22.Text;
            d[24] = txtOutput31.Text;
            d[25] = txtOutput27.Text;
            d[26] = txtOutput28.Text;
            for (int i = 0; i < 13; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(2).SetCellValue(d[i]);
            }
            for (int i = 0; i < 14; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(7).SetCellValue(d[13 + i]);
            }
            #endregion
            #region   投资/效益导出
            string[] c = new string[8];
            c[0] = txtOutput1.Text;
            c[1] = txtOutput2.Text;
            c[2] = txtOutput4.Text;

            c[3] = txtOutput5.Text;
            c[4] = txtOutput6.Text;
            c[5] = txtOutput8.Text;
            c[6] = txtOutput9.Text;
            c[7] = txtOutput10.Text;

            sheet1[0].GetRow(20).GetCell(2).SetCellValue(c[0]);
            sheet1[0].GetRow(21).GetCell(2).SetCellValue(c[1]);
            sheet1[0].GetRow(22).GetCell(2).SetCellValue(c[2]);
            for (int i = 0; i < 5; i++)
            {
                sheet1[0].GetRow(20 + i).GetCell(7).SetCellValue(c[3 + i]);
            }
            #endregion
            #region 收入/成本导出
            string[] b = new string[21];
            b[0] = txtOutput11.Text;
            b[1] = txtOutput12.Text;
            b[2] = txtOutput13.Text;
            b[3] = txtOutput14.Text;
            b[4] = txtOutput15.Text;
            b[5] = txtOutput16.Text;
            b[6] = txtOutput17.Text;
            b[7] = txtOutput18.Text;
            b[8] = txtOutput19.Text;
            b[9] = txtOutput20.Text;

            b[10] = txtOutput21.Text;
            b[11] = txtOutput22.Text;
            b[12] = txtOutput23.Text;
            b[13] = txtOutput24.Text;
            b[14] = txtOutput25.Text;
            b[15] = txtOutput26.Text;
            b[16] = txtOutput29.Text;
            b[17] = txtOutput30.Text;
            b[18] = txtOutput32.Text;
            b[19] = txtOutput33.Text;
            b[20] = txtOutput34.Text;

            for (int i = 0; i < 9; i++)
            {
                sheet1[0].GetRow(28 + i).GetCell(2).SetCellValue(b[i]);
            }
            for (int i = 10; i < 20; i++)
            {
                sheet1[0].GetRow(18 + i).GetCell(7).SetCellValue(b[i]);
            }
            #endregion
            #region  把datagridview中的数据导出到Excel
            string[] a = new string[32];
            for (int i = 0; i < 9; i++)
            {
                for (int j = 0; j < 32; j++)
                {
                    a[j] = (dataGridView1.Rows[j].Cells[i].Value).ToString();
                    sheet1[0].GetRow(43 + i).GetCell(4 + j).SetCellValue(a[j]);
                }
            }
            #endregion

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "分布式能源投资决策模型--技术经济型模型.cs";//设置默认文件名
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

        private void button16_Click(object sender, EventArgs e)
        {
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求分析--分布式能源1.xlsx";
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

            #region   输入导出
            string[] d = new string[27];
            d[0] = comboBox1.Text;
            d[1] = txtInput2.Text;
            d[2] = txtInput3.Text;
            d[3] = txtInput4.Text;
            d[4] = txtInput5.Text;
            d[5] = txtInput6.Text;
            d[6] = txtInput7.Text;
            d[7] = txtInput8.Text;
            d[8] = txtInput9.Text;
            d[9] = txtInput10.Text;
            d[10] = txtInput11.Text;
            d[11] = txtOutput3.Text;
            d[12] = txtOutput7.Text;

            d[13] = txtInput12.Text;
            d[14] = txtInput13.Text;
            d[15] = txtInput14.Text;
            d[16] = txtInput15.Text;
            d[17] = txtInput16.Text;
            d[18] = txtInput17.Text;
            d[19] = txtInput18.Text;
            d[20] = txtInput19.Text;
            d[21] = txtInput20.Text;
            d[22] = txtInput21.Text;
            d[23] = txtInput22.Text;
            d[24] = txtOutput31.Text;
            d[25] = txtOutput27.Text;
            d[26] = txtOutput28.Text;
            for (int i = 0; i < 13; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(2).SetCellValue(d[i]);
            }
            for (int i = 0; i < 14; i++)
            {
                sheet1[0].GetRow(4 + i).GetCell(7).SetCellValue(d[13 + i]);
            }
            #endregion
            #region   投资/效益导出
            string[] c = new string[8];
            c[0] = txtOutput1.Text;
            c[1] = txtOutput2.Text;
            c[2] = txtOutput4.Text;

            c[3] = txtOutput5.Text;
            c[4] = txtOutput6.Text;
            c[5] = txtOutput8.Text;
            c[6] = txtOutput9.Text;
            c[7] = txtOutput10.Text;

            sheet1[0].GetRow(20).GetCell(2).SetCellValue(c[0]);
            sheet1[0].GetRow(21).GetCell(2).SetCellValue(c[1]);
            sheet1[0].GetRow(22).GetCell(2).SetCellValue(c[2]);
            for (int i = 0; i < 5; i++)
            {
                sheet1[0].GetRow(20 + i).GetCell(7).SetCellValue(c[3 + i]);
            }
            #endregion
            #region 收入/成本导出
            string[] b = new string[21];
            b[0] = txtOutput11.Text;
            b[1] = txtOutput12.Text;
            b[2] = txtOutput13.Text;
            b[3] = txtOutput14.Text;
            b[4] = txtOutput15.Text;
            b[5] = txtOutput16.Text;
            b[6] = txtOutput17.Text;
            b[7] = txtOutput18.Text;
            b[8] = txtOutput19.Text;
            b[9] = txtOutput20.Text;

            b[10] = txtOutput21.Text;
            b[11] = txtOutput22.Text;
            b[12] = txtOutput23.Text;
            b[13] = txtOutput24.Text;
            b[14] = txtOutput25.Text;
            b[15] = txtOutput26.Text;
            b[16] = txtOutput29.Text;
            b[17] = txtOutput30.Text;
            b[18] = txtOutput32.Text;
            b[19] = txtOutput33.Text;
            b[20] = txtOutput34.Text;

            for (int i = 0; i < 9; i++)
            {
                sheet1[0].GetRow(28 + i).GetCell(2).SetCellValue(b[i]);
            }
            for (int i = 10; i < 20; i++)
            {
                sheet1[0].GetRow(18 + i).GetCell(7).SetCellValue(b[i]);
            }
            #endregion
            #region  把datagridview中的数据导出到Excel
            string[] a = new string[32];
            for (int i = 0; i < 9; i++)
            {
                for (int j = 0; j < 32; j++)
                {
                    a[j] = (dataGridView1.Rows[j].Cells[i].Value).ToString();
                    sheet1[0].GetRow(43 + i).GetCell(4 + j).SetCellValue(a[j]);
                }
            }
            #endregion

            string path = null;
            saveFile.Filter = "xlsx(*.xlsx)|*.xlsx";//设置文件类型
            saveFile.FileName = "分布式能源投资决策模型--技术经济型模型.cs";//设置默认文件名
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






