using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class Windows7 : Form
    {
        public Windows7()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        #region  界面datagridview填充
        private void Windows7_Load(object sender, EventArgs e)
        {
            #region Datagridview1中数据加载
            dataGridView1.EnableHeadersVisualStyles = false;// 变灰
            dataGridView1.TopLeftHeaderCell.Value = "年份";
            int index = this.dataGridView1.Rows.Add(21);
            double k = 2010;
            for (int i = 0; i < 21; i++)
            {
                this.dataGridView1.Rows[i].HeaderCell.Value = Convert.ToString(k);
                k++;
            }
            this.dataGridView1.RowHeadersWidth = 61;//设置宽度
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView1.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;


            this.dataGridView1.Rows[0].Cells[0].Value = 3410;
            this.dataGridView1.Rows[1].Cells[0].Value = 3440;
            this.dataGridView1.Rows[2].Cells[0].Value = 3466;
            this.dataGridView1.Rows[3].Cells[0].Value = 3488;
            this.dataGridView1.Rows[4].Cells[0].Value = 3511;
            this.dataGridView1.Rows[5].Cells[0].Value = 3535;
            this.dataGridView1.Rows[6].Cells[0].Value = 3558;
            this.dataGridView1.Rows[7].Cells[0].Value = 3581;
            this.dataGridView1.Rows[8].Cells[0].Value = 3604;
            this.dataGridView1.Rows[9].Cells[0].Value = 3627;
            this.dataGridView1.Rows[10].Cells[0].Value = 3693;
            this.dataGridView1.Rows[11].Cells[0].Value = 0;
            this.dataGridView1.Rows[12].Cells[0].Value = 0;
            this.dataGridView1.Rows[13].Cells[0].Value = 0;
            this.dataGridView1.Rows[14].Cells[0].Value = 0;
            this.dataGridView1.Rows[15].Cells[0].Value = 0;
            this.dataGridView1.Rows[16].Cells[0].Value = 0;
            this.dataGridView1.Rows[17].Cells[0].Value = 0;
            this.dataGridView1.Rows[18].Cells[0].Value = 0;
            this.dataGridView1.Rows[19].Cells[0].Value = 0;
            this.dataGridView1.Rows[20].Cells[0].Value = 0;

            this.dataGridView1.Rows[0].Cells[3].Value = 0;
            this.dataGridView1.Rows[0].Cells[13].Value = 0;
            this.dataGridView1.Rows[0].Cells[15].Value = 0;
            this.dataGridView1.Rows[0].Cells[18].Value = 0;
            this.dataGridView1.Rows[0].Cells[20].Value = 0;


            this.dataGridView1.Rows[0].Cells[1].Value = 5.8;
            this.dataGridView1.Rows[1].Cells[1].Value = "6.0";
            this.dataGridView1.Rows[2].Cells[1].Value = 5.8;
            this.dataGridView1.Rows[3].Cells[1].Value = 5.9;
            this.dataGridView1.Rows[4].Cells[1].Value = "6.0";
            this.dataGridView1.Rows[5].Cells[1].Value = "6.0";
            this.dataGridView1.Rows[6].Cells[1].Value = 6.3;
            this.dataGridView1.Rows[7].Cells[1].Value = 6.1;
            this.dataGridView1.Rows[8].Cells[1].Value = "6.0";
            this.dataGridView1.Rows[9].Cells[1].Value = "6.0";
            this.dataGridView1.Rows[10].Cells[1].Value = 6.1;
            this.dataGridView1.Rows[11].Cells[1].Value = 0;
            this.dataGridView1.Rows[12].Cells[1].Value = 0;
            this.dataGridView1.Rows[13].Cells[1].Value = 0;
            this.dataGridView1.Rows[14].Cells[1].Value = 0;
            this.dataGridView1.Rows[15].Cells[1].Value = 0;
            this.dataGridView1.Rows[16].Cells[1].Value = 0;
            this.dataGridView1.Rows[17].Cells[1].Value = 0;
            this.dataGridView1.Rows[18].Cells[1].Value = 0;
            this.dataGridView1.Rows[19].Cells[1].Value = 0;
            this.dataGridView1.Rows[20].Cells[1].Value = 0;


            this.dataGridView1.Rows[0].Cells[2].Value = 1431;
            this.dataGridView1.Rows[1].Cells[2].Value = 1462;
            this.dataGridView1.Rows[2].Cells[2].Value = 1546;
            this.dataGridView1.Rows[3].Cells[2].Value = 1573;
            this.dataGridView1.Rows[4].Cells[2].Value = 1615;
            this.dataGridView1.Rows[5].Cells[2].Value = 1672;
            this.dataGridView1.Rows[6].Cells[2].Value = 1708;
            this.dataGridView1.Rows[7].Cells[2].Value = 1744;
            this.dataGridView1.Rows[8].Cells[2].Value = 1798;
            this.dataGridView1.Rows[9].Cells[2].Value = 1864;
            this.dataGridView1.Rows[10].Cells[2].Value = 2108;
            this.dataGridView1.Rows[11].Cells[2].Value = 0;
            this.dataGridView1.Rows[12].Cells[2].Value = 0;
            this.dataGridView1.Rows[13].Cells[2].Value = 0;
            this.dataGridView1.Rows[14].Cells[2].Value = 0;
            this.dataGridView1.Rows[15].Cells[2].Value = 0;
            this.dataGridView1.Rows[16].Cells[2].Value = 0;
            this.dataGridView1.Rows[17].Cells[2].Value = 0;
            this.dataGridView1.Rows[18].Cells[2].Value = 0;
            this.dataGridView1.Rows[19].Cells[2].Value = 0;
            this.dataGridView1.Rows[20].Cells[2].Value = 0;

            this.dataGridView1.Rows[0].Cells[5].Value = 3765;
            this.dataGridView1.Rows[1].Cells[5].Value = 4073;
            this.dataGridView1.Rows[2].Cells[5].Value = 4468;
            this.dataGridView1.Rows[3].Cells[5].Value = 4984;
            this.dataGridView1.Rows[4].Cells[5].Value = 5763;
            this.dataGridView1.Rows[5].Cells[5].Value = 6569;
            this.dataGridView1.Rows[6].Cells[5].Value = 7584;
            this.dataGridView1.Rows[7].Cells[5].Value = 9249;
            this.dataGridView1.Rows[8].Cells[5].Value = 10823;
            this.dataGridView1.Rows[9].Cells[5].Value = 12236;
            this.dataGridView1.Rows[10].Cells[5].Value = 14737;
            this.dataGridView1.Rows[11].Cells[5].Value = 0;
            this.dataGridView1.Rows[12].Cells[5].Value = 0;
            this.dataGridView1.Rows[13].Cells[5].Value = 0;
            this.dataGridView1.Rows[14].Cells[5].Value = 0;
            this.dataGridView1.Rows[15].Cells[5].Value = 0;
            this.dataGridView1.Rows[16].Cells[5].Value = 0;
            this.dataGridView1.Rows[17].Cells[5].Value = 0;
            this.dataGridView1.Rows[18].Cells[5].Value = 0;
            this.dataGridView1.Rows[19].Cells[5].Value = 0;
            this.dataGridView1.Rows[20].Cells[5].Value = 0;

            this.dataGridView1.Rows[0].Cells[6].Value = 109.3;
            this.dataGridView1.Rows[1].Cells[6].Value = 108.7;
            this.dataGridView1.Rows[2].Cells[6].Value = 110.2;
            this.dataGridView1.Rows[3].Cells[6].Value = 111.5;
            this.dataGridView1.Rows[4].Cells[6].Value = 111.8;
            this.dataGridView1.Rows[5].Cells[6].Value = 111.6;
            this.dataGridView1.Rows[6].Cells[6].Value = 114.8;
            this.dataGridView1.Rows[7].Cells[6].Value = 115.2;
            this.dataGridView1.Rows[8].Cells[6].Value = 113.0;
            this.dataGridView1.Rows[9].Cells[6].Value = 112.3;
            this.dataGridView1.Rows[10].Cells[6].Value = 113.9;
            this.dataGridView1.Rows[11].Cells[6].Value = 0;
            this.dataGridView1.Rows[12].Cells[6].Value = 0;
            this.dataGridView1.Rows[13].Cells[6].Value = 0;
            this.dataGridView1.Rows[14].Cells[6].Value = 0;
            this.dataGridView1.Rows[15].Cells[6].Value = 0;
            this.dataGridView1.Rows[16].Cells[6].Value = 0;
            this.dataGridView1.Rows[17].Cells[6].Value = 0;
            this.dataGridView1.Rows[18].Cells[6].Value = 0;
            this.dataGridView1.Rows[19].Cells[6].Value = 0;
            this.dataGridView1.Rows[20].Cells[6].Value = 0;

            this.dataGridView1.Rows[0].Cells[8].Value = 3463;
            this.dataGridView1.Rows[1].Cells[8].Value = 3163;
            this.dataGridView1.Rows[2].Cells[8].Value = 3490;
            this.dataGridView1.Rows[3].Cells[8].Value = 3925;
            this.dataGridView1.Rows[4].Cells[8].Value = 5449;
            this.dataGridView1.Rows[5].Cells[8].Value = 6157;
            this.dataGridView1.Rows[6].Cells[8].Value = 6840;
            this.dataGridView1.Rows[7].Cells[8].Value = 7574;
            this.dataGridView1.Rows[8].Cells[8].Value = 8254;
            this.dataGridView1.Rows[9].Cells[8].Value = 8917;
            this.dataGridView1.Rows[10].Cells[8].Value = 9809;
            this.dataGridView1.Rows[11].Cells[8].Value = 0;
            this.dataGridView1.Rows[12].Cells[8].Value = 0;
            this.dataGridView1.Rows[13].Cells[8].Value = 0;
            this.dataGridView1.Rows[14].Cells[8].Value = 0;
            this.dataGridView1.Rows[15].Cells[8].Value = 0;
            this.dataGridView1.Rows[16].Cells[8].Value = 0;
            this.dataGridView1.Rows[17].Cells[8].Value = 0;
            this.dataGridView1.Rows[18].Cells[8].Value = 0;
            this.dataGridView1.Rows[19].Cells[8].Value = 0;
            this.dataGridView1.Rows[20].Cells[8].Value = 0;

            this.dataGridView1.Rows[0].Cells[9].Value = 54.4;
            this.dataGridView1.Rows[1].Cells[9].Value = 51.4;
            this.dataGridView1.Rows[2].Cells[9].Value = 55.6;
            this.dataGridView1.Rows[3].Cells[9].Value = 61.4;
            this.dataGridView1.Rows[4].Cells[9].Value = 63.8;
            this.dataGridView1.Rows[5].Cells[9].Value = 62;
            this.dataGridView1.Rows[6].Cells[9].Value = 61.8;
            this.dataGridView1.Rows[7].Cells[9].Value = 65.2;
            this.dataGridView1.Rows[8].Cells[9].Value = 64.8;
            this.dataGridView1.Rows[9].Cells[9].Value = 67.6;
            this.dataGridView1.Rows[10].Cells[9].Value = 57.8;
            this.dataGridView1.Rows[11].Cells[9].Value = 0;
            this.dataGridView1.Rows[12].Cells[9].Value = 0;
            this.dataGridView1.Rows[13].Cells[9].Value = 0;
            this.dataGridView1.Rows[14].Cells[9].Value = 0;
            this.dataGridView1.Rows[15].Cells[9].Value = 0;
            this.dataGridView1.Rows[16].Cells[9].Value = 0;
            this.dataGridView1.Rows[17].Cells[9].Value = 0;
            this.dataGridView1.Rows[18].Cells[9].Value = 0;
            this.dataGridView1.Rows[19].Cells[9].Value = 0;
            this.dataGridView1.Rows[20].Cells[9].Value = 0;

            this.dataGridView1.Rows[0].Cells[10].Value = 23.3;
            this.dataGridView1.Rows[1].Cells[10].Value = 22.0;
            this.dataGridView1.Rows[2].Cells[10].Value = 23.8;
            this.dataGridView1.Rows[3].Cells[10].Value = 24.5;
            this.dataGridView1.Rows[4].Cells[10].Value = 25.1;
            this.dataGridView1.Rows[5].Cells[10].Value = 22.3;
            this.dataGridView1.Rows[6].Cells[10].Value = 21.6;
            this.dataGridView1.Rows[7].Cells[10].Value = 21.4;
            this.dataGridView1.Rows[8].Cells[10].Value = 18.9;
            this.dataGridView1.Rows[9].Cells[10].Value = 18.3;
            this.dataGridView1.Rows[10].Cells[10].Value = 23.6;
            this.dataGridView1.Rows[11].Cells[10].Value = 0;
            this.dataGridView1.Rows[12].Cells[10].Value = 0;
            this.dataGridView1.Rows[13].Cells[10].Value = 0;
            this.dataGridView1.Rows[14].Cells[10].Value = 0;
            this.dataGridView1.Rows[15].Cells[10].Value = 0;
            this.dataGridView1.Rows[16].Cells[10].Value = 0;
            this.dataGridView1.Rows[17].Cells[10].Value = 0;
            this.dataGridView1.Rows[18].Cells[10].Value = 0;
            this.dataGridView1.Rows[19].Cells[10].Value = 0;
            this.dataGridView1.Rows[20].Cells[10].Value = 0;

            this.dataGridView1.Rows[0].Cells[11].Value = 0;
            this.dataGridView1.Rows[1].Cells[11].Value = 0;
            this.dataGridView1.Rows[2].Cells[11].Value = 0;
            this.dataGridView1.Rows[3].Cells[11].Value = 0;
            this.dataGridView1.Rows[4].Cells[11].Value = 0.2;
            this.dataGridView1.Rows[5].Cells[11].Value = 0.1;
            this.dataGridView1.Rows[6].Cells[11].Value = 0.1;
            this.dataGridView1.Rows[7].Cells[11].Value = 0.1;
            this.dataGridView1.Rows[8].Cells[11].Value = 0.3;
            this.dataGridView1.Rows[9].Cells[11].Value = 1.3;
            this.dataGridView1.Rows[10].Cells[11].Value = 4.0;
            this.dataGridView1.Rows[11].Cells[11].Value = 0;
            this.dataGridView1.Rows[12].Cells[11].Value = 0;
            this.dataGridView1.Rows[13].Cells[11].Value = 0;
            this.dataGridView1.Rows[14].Cells[11].Value = 0;
            this.dataGridView1.Rows[15].Cells[11].Value = 0;
            this.dataGridView1.Rows[16].Cells[11].Value = 0;
            this.dataGridView1.Rows[17].Cells[11].Value = 0;
            this.dataGridView1.Rows[18].Cells[11].Value = 0;
            this.dataGridView1.Rows[19].Cells[11].Value = 0;
            this.dataGridView1.Rows[20].Cells[11].Value = 0;


            this.dataGridView1.Rows[0].Cells[12].Value = 22.3;
            this.dataGridView1.Rows[1].Cells[12].Value = 26.6;
            this.dataGridView1.Rows[2].Cells[12].Value = 20.6;
            this.dataGridView1.Rows[3].Cells[12].Value = 14.1;
            this.dataGridView1.Rows[4].Cells[12].Value = 10.9;
            this.dataGridView1.Rows[5].Cells[12].Value = 15.6;
            this.dataGridView1.Rows[6].Cells[12].Value = 16.5;
            this.dataGridView1.Rows[7].Cells[12].Value = 13.3;
            this.dataGridView1.Rows[8].Cells[12].Value = 16;
            this.dataGridView1.Rows[9].Cells[12].Value = 12.8;
            this.dataGridView1.Rows[10].Cells[12].Value = 14.6;
            this.dataGridView1.Rows[11].Cells[12].Value = 0;
            this.dataGridView1.Rows[12].Cells[12].Value = 0;
            this.dataGridView1.Rows[13].Cells[12].Value = 0;
            this.dataGridView1.Rows[14].Cells[12].Value = 0;
            this.dataGridView1.Rows[15].Cells[12].Value = 0;
            this.dataGridView1.Rows[16].Cells[12].Value = 0;
            this.dataGridView1.Rows[17].Cells[12].Value = 0;
            this.dataGridView1.Rows[18].Cells[12].Value = 0;
            this.dataGridView1.Rows[19].Cells[12].Value = 0;
            this.dataGridView1.Rows[20].Cells[12].Value = 0;


            this.dataGridView1.Rows[0].Cells[14].Value = 0.66;
            this.dataGridView1.Rows[1].Cells[14].Value = 0.86;
            this.dataGridView1.Rows[2].Cells[14].Value = 1.40;
            this.dataGridView1.Rows[3].Cells[14].Value = 1.08;
            this.dataGridView1.Rows[4].Cells[14].Value = 0.97;
            this.dataGridView1.Rows[5].Cells[14].Value = 1.12;
            this.dataGridView1.Rows[6].Cells[14].Value = 0.75;
            this.dataGridView1.Rows[7].Cells[14].Value = 0.73;
            this.dataGridView1.Rows[8].Cells[14].Value = 0.68;
            this.dataGridView1.Rows[9].Cells[14].Value = 0.65;
            this.dataGridView1.Rows[10].Cells[14].Value = 0.72;
            this.dataGridView1.Rows[11].Cells[14].Value = 0;
            this.dataGridView1.Rows[12].Cells[14].Value = 0;
            this.dataGridView1.Rows[13].Cells[14].Value = 0;
            this.dataGridView1.Rows[14].Cells[14].Value = 0;
            this.dataGridView1.Rows[15].Cells[14].Value = 0;
            this.dataGridView1.Rows[16].Cells[14].Value = 0;
            this.dataGridView1.Rows[17].Cells[14].Value = 0;
            this.dataGridView1.Rows[18].Cells[14].Value = 0;
            this.dataGridView1.Rows[19].Cells[14].Value = 0;
            this.dataGridView1.Rows[20].Cells[14].Value = 0;


            this.dataGridView1.Rows[0].Cells[16].Value = 0;
            this.dataGridView1.Rows[1].Cells[16].Value = 0;
            this.dataGridView1.Rows[2].Cells[16].Value = 0;
            this.dataGridView1.Rows[3].Cells[16].Value = 0;
            this.dataGridView1.Rows[4].Cells[16].Value = 0;
            this.dataGridView1.Rows[5].Cells[16].Value = 0.94;
            this.dataGridView1.Rows[6].Cells[16].Value = 0.91;
            this.dataGridView1.Rows[7].Cells[16].Value = 0.88;
            this.dataGridView1.Rows[8].Cells[16].Value = 0.85;
            this.dataGridView1.Rows[9].Cells[16].Value = 0.81;
            this.dataGridView1.Rows[10].Cells[16].Value = 0.78;
            this.dataGridView1.Rows[11].Cells[16].Value = 0;
            this.dataGridView1.Rows[12].Cells[16].Value = 0;
            this.dataGridView1.Rows[13].Cells[16].Value = 0;
            this.dataGridView1.Rows[14].Cells[16].Value = 0;
            this.dataGridView1.Rows[15].Cells[16].Value = 0;
            this.dataGridView1.Rows[16].Cells[16].Value = 0;
            this.dataGridView1.Rows[17].Cells[16].Value = 0;
            this.dataGridView1.Rows[18].Cells[16].Value = 0;
            this.dataGridView1.Rows[19].Cells[16].Value = 0;
            this.dataGridView1.Rows[20].Cells[16].Value = 0;

            this.dataGridView1.Rows[0].Cells[17].Value = 403;
            this.dataGridView1.Rows[1].Cells[17].Value = 439;
            this.dataGridView1.Rows[2].Cells[17].Value = 498;
            this.dataGridView1.Rows[3].Cells[17].Value = 585;
            this.dataGridView1.Rows[4].Cells[17].Value = 665;
            this.dataGridView1.Rows[5].Cells[17].Value = 757;
            this.dataGridView1.Rows[6].Cells[17].Value = 867;
            this.dataGridView1.Rows[7].Cells[17].Value = 1000;
            this.dataGridView1.Rows[8].Cells[17].Value = 1074;
            this.dataGridView1.Rows[9].Cells[17].Value = 1135;
            this.dataGridView1.Rows[10].Cells[17].Value = 1315;
            this.dataGridView1.Rows[11].Cells[17].Value = 0;
            this.dataGridView1.Rows[12].Cells[17].Value = 0;
            this.dataGridView1.Rows[13].Cells[17].Value = 0;
            this.dataGridView1.Rows[14].Cells[17].Value = 0;
            this.dataGridView1.Rows[15].Cells[17].Value = 0;
            this.dataGridView1.Rows[16].Cells[17].Value = 0;
            this.dataGridView1.Rows[17].Cells[17].Value = 0;
            this.dataGridView1.Rows[18].Cells[17].Value = 0;
            this.dataGridView1.Rows[19].Cells[17].Value = 0;
            this.dataGridView1.Rows[20].Cells[17].Value = 0;


            this.dataGridView1.Rows[0].Cells[19].Value = 1.45;
            this.dataGridView1.Rows[1].Cells[19].Value = 1.05;
            this.dataGridView1.Rows[2].Cells[19].Value = 2.13;
            this.dataGridView1.Rows[3].Cells[19].Value = 1.54;
            this.dataGridView1.Rows[4].Cells[19].Value = 0.45;
            this.dataGridView1.Rows[5].Cells[19].Value = 1.20;
            this.dataGridView1.Rows[6].Cells[19].Value = 0.98;
            this.dataGridView1.Rows[7].Cells[19].Value = 1.02;
            this.dataGridView1.Rows[8].Cells[19].Value = 0.56;
            this.dataGridView1.Rows[9].Cells[19].Value = 0.47;
            this.dataGridView1.Rows[10].Cells[19].Value = 1.14;
            this.dataGridView1.Rows[11].Cells[19].Value = 0;
            this.dataGridView1.Rows[12].Cells[19].Value = 0;
            this.dataGridView1.Rows[13].Cells[19].Value = 0;
            this.dataGridView1.Rows[14].Cells[19].Value = 0;
            this.dataGridView1.Rows[15].Cells[19].Value = 0;
            this.dataGridView1.Rows[16].Cells[19].Value = 0;
            this.dataGridView1.Rows[17].Cells[19].Value = 0;
            this.dataGridView1.Rows[18].Cells[19].Value = 0;
            this.dataGridView1.Rows[19].Cells[19].Value = 0;
            this.dataGridView1.Rows[20].Cells[19].Value = 0;



            #endregion

            #region  datagridview2中数据加载
            dataGridView2.EnableHeadersVisualStyles = false;// 变灰
            dataGridView2.TopLeftHeaderCell.Value = "年份";
            int index1 = this.dataGridView2.Rows.Add(8);
            double k1 = 2015;
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView2.Rows[i].HeaderCell.Value = Convert.ToString(k1);
                k1 += 5;
            }
            this.dataGridView2.RowHeadersWidth = 61;//设置宽度
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView2.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            this.dataGridView2.Rows[1].Cells[0].Value = 0.41;
            this.dataGridView2.Rows[2].Cells[0].Value = 0.45;
            this.dataGridView2.Rows[3].Cells[0].Value = 0.48;
            this.dataGridView2.Rows[4].Cells[0].Value = 0.51;

            this.dataGridView2.Rows[1].Cells[1].Value = 0.39;
            this.dataGridView2.Rows[2].Cells[1].Value = 0.43;
            this.dataGridView2.Rows[3].Cells[1].Value = 0.46;
            this.dataGridView2.Rows[4].Cells[1].Value = 0.49;

            this.dataGridView2.Rows[1].Cells[2].Value = 0.37;
            this.dataGridView2.Rows[2].Cells[2].Value = 0.41;
            this.dataGridView2.Rows[3].Cells[2].Value = 0.44;
            this.dataGridView2.Rows[4].Cells[2].Value = 0.47;
            for (int i = 5; i < 8; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    this.dataGridView2.Rows[i].Cells[j].Value = 0;
                }

            }
            #endregion

            #region  Datagridview的设置
            dataGridView3.EnableHeadersVisualStyles = false;// 变灰
            dataGridView3.TopLeftHeaderCell.Value = "年份";
            int index2 = this.dataGridView3.Rows.Add(8);

            dataGridView4.EnableHeadersVisualStyles = false;// 变灰
            dataGridView4.TopLeftHeaderCell.Value = "年份";
            int index3 = this.dataGridView4.Rows.Add(8);
            this.dataGridView4.RowHeadersWidth = 64;//设置宽度
            dataGridView4.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView4.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView4.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView5.EnableHeadersVisualStyles = false;// 变灰
            dataGridView5.TopLeftHeaderCell.Value = "年份";
            int index4 = this.dataGridView5.Rows.Add(8);
            this.dataGridView5.RowHeadersWidth = 64;//设置宽度
            dataGridView5.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView5.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView5.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView6.EnableHeadersVisualStyles = false;// 变灰
            dataGridView6.TopLeftHeaderCell.Value = "年份";
            int index5 = this.dataGridView6.Rows.Add(8);
            this.dataGridView6.RowHeadersWidth = 64;//设置宽度
            dataGridView6.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView6.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView6.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView7.EnableHeadersVisualStyles = false;// 变灰
            dataGridView7.TopLeftHeaderCell.Value = "年份";
            int index6 = this.dataGridView7.Rows.Add(8);
            this.dataGridView7.RowHeadersWidth = 64;//设置宽度
            dataGridView7.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView7.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView7.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView8.EnableHeadersVisualStyles = false;// 变灰
            dataGridView8.TopLeftHeaderCell.Value = "年份";
            int index7 = this.dataGridView8.Rows.Add(8);
            this.dataGridView8.RowHeadersWidth = 64;//设置宽度
            dataGridView8.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView8.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView8.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView9.EnableHeadersVisualStyles = false;// 变灰
            dataGridView9.TopLeftHeaderCell.Value = "序号";
            int index8 = this.dataGridView9.Rows.Add(12);
            this.dataGridView9.RowHeadersWidth = 64;//设置宽度
            dataGridView9.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView9.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView9.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            this.dataGridView9.Rows[0].HeaderCell.Value = Convert.ToString(1);
            this.dataGridView9.Rows[1].HeaderCell.Value = Convert.ToString(2);

            double k2 = 2015;
            for (int i = 0; i < 8; i++)
            {
                this.dataGridView3.Rows[i].HeaderCell.Value = Convert.ToString(k2);
                this.dataGridView4.Rows[i].HeaderCell.Value = Convert.ToString(k2);
                this.dataGridView5.Rows[i].HeaderCell.Value = Convert.ToString(k2);
                this.dataGridView6.Rows[i].HeaderCell.Value = Convert.ToString(k2);
                this.dataGridView7.Rows[i].HeaderCell.Value = Convert.ToString(k2);
                this.dataGridView8.Rows[i].HeaderCell.Value = Convert.ToString(k2);
                k2 += 5;
            }
            this.dataGridView3.RowHeadersWidth = 61;//设置宽度
            dataGridView3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView3.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;//设置列宽度是否可变
            dataGridView3.TopLeftHeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            #endregion

            #region Datagridview3中数据的加载

            this.dataGridView3.Rows[1].Cells[0].Value = "9.0";
            this.dataGridView3.Rows[2].Cells[0].Value = "8.0";
            this.dataGridView3.Rows[3].Cells[0].Value = "7.0";
            this.dataGridView3.Rows[4].Cells[0].Value = 6.5;

            this.dataGridView3.Rows[1].Cells[1].Value = 8.5;
            this.dataGridView3.Rows[2].Cells[1].Value = 7.5;
            this.dataGridView3.Rows[3].Cells[1].Value = 6.5;
            this.dataGridView3.Rows[4].Cells[1].Value = "6.0";

            this.dataGridView3.Rows[1].Cells[2].Value = "8.0";
            this.dataGridView3.Rows[2].Cells[2].Value = "7.0";
            this.dataGridView3.Rows[3].Cells[2].Value = "6.0";
            this.dataGridView3.Rows[4].Cells[2].Value = 5.5;
            for (int i = 5; i < 8; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    this.dataGridView3.Rows[i].Cells[j].Value = 0;
                }



                #endregion

            #region Datagidview4中数据的加载
                this.dataGridView4.Rows[1].Cells[0].Value = "7.0";
                this.dataGridView4.Rows[2].Cells[0].Value = "6.0";
                this.dataGridView4.Rows[3].Cells[0].Value = "4.0";
                this.dataGridView4.Rows[4].Cells[0].Value = "3.0";

                this.dataGridView4.Rows[1].Cells[1].Value = "6.0";
                this.dataGridView4.Rows[2].Cells[1].Value = "4.0";
                this.dataGridView4.Rows[3].Cells[1].Value = "3.0";
                this.dataGridView4.Rows[4].Cells[1].Value = "2.0";

                this.dataGridView4.Rows[1].Cells[2].Value = "5.0";
                this.dataGridView4.Rows[2].Cells[2].Value = "4.0";
                this.dataGridView4.Rows[3].Cells[2].Value = "3.0";
                this.dataGridView4.Rows[4].Cells[2].Value = "2.0";
                for (int m = 5; m < 8; m++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        this.dataGridView4.Rows[m].Cells[j].Value = 0;
                    }

                    #endregion

            #region Datagridvirew5中数据的填充
                    this.dataGridView5.Rows[0].Cells[0].Value = "3.4";
                    this.dataGridView5.Rows[1].Cells[0].Value = "4.0";
                    this.dataGridView5.Rows[2].Cells[0].Value = "3.0";
                    this.dataGridView5.Rows[3].Cells[0].Value = "2.0";
                    this.dataGridView5.Rows[4].Cells[0].Value = "1.8";

                    this.dataGridView5.Rows[0].Cells[1].Value = "3.4";
                    this.dataGridView5.Rows[1].Cells[1].Value = "4.4";
                    this.dataGridView5.Rows[2].Cells[1].Value = "3.4";
                    this.dataGridView5.Rows[3].Cells[1].Value = "2.2";
                    this.dataGridView5.Rows[4].Cells[1].Value = "2.0";

                    this.dataGridView5.Rows[0].Cells[2].Value = "3.4";
                    this.dataGridView5.Rows[1].Cells[2].Value = "5.0";
                    this.dataGridView5.Rows[2].Cells[2].Value = "4.0";
                    this.dataGridView5.Rows[3].Cells[2].Value = "2.5";
                    this.dataGridView5.Rows[4].Cells[2].Value = "2.2";

                    this.dataGridView5.Rows[0].Cells[3].Value = "3.4";
                    this.dataGridView5.Rows[1].Cells[3].Value = "7.0";
                    this.dataGridView5.Rows[2].Cells[3].Value = "8.0";
                    this.dataGridView5.Rows[3].Cells[3].Value = "9.0";
                    this.dataGridView5.Rows[4].Cells[3].Value = "10.0";

                    for (int n = 5; n < 8; n++)
                    {
                        for (int j = 0; j < 4; j++)
                        {
                            this.dataGridView5.Rows[n].Cells[j].Value = 0;
                        }
                    }

                    #endregion

                    #region  Datagridview6、7、8中数据导入加载

                    this.dataGridView6.Rows[0].Cells[7].Value = "27.6";
                    this.dataGridView6.Rows[1].Cells[7].Value = "74.8";
                    this.dataGridView6.Rows[2].Cells[7].Value = "106.1";

                    this.dataGridView7.Rows[0].Cells[7].Value = "27.6";
                    this.dataGridView7.Rows[1].Cells[7].Value = "74.8";
                    this.dataGridView7.Rows[2].Cells[7].Value = "106.1";

                    this.dataGridView8.Rows[0].Cells[7].Value = "27.6";
                    this.dataGridView8.Rows[1].Cells[7].Value = "74.8";
                    this.dataGridView8.Rows[2].Cells[7].Value = "106.1";










                    #endregion
                }
            }
        }
        #endregion



        #region 能耗系数法计算
        private void button5_Click(object sender, EventArgs e)
        {
            Calculate();
        }
        
        #region     城镇人口预测结果计算

        private double Easy(double K1, double K2)
        {

            int T1 = Convert.ToInt32(textBox2.Text);
            double S = 0;
            for (int i = 0; i < 5; i++)
            {
                S = S + Convert.ToDouble(this.dataGridView1.Rows[T1-2010 - i].Cells[1].Value);
            }
            double T2 = K1 - 2010;

            double T3 = Math.Pow(1 + S / 3 / 1000, T2);

            double T4 = Convert.ToDouble(this.dataGridView1.Rows[T1-2010 - 4].Cells[0].Value) * T3 * K2;
            return T4;
        }
        #endregion

        #region  增长法算法
        private double Easy1(double K1, double K2)
        {
            double T1 = 5;
            double T2 = Math.Pow((1 + K2 / 100), T1);
            double T3 = K1 * T2;
            return T3;
        }
        #endregion

        #region  能耗法算法
        private double Easy2(double K1,double K2,double K3)
        {
            double T1 = Math.Pow((1 + K2/100), 5);
            double T2 = Math.Pow((1 - K3/100), 5);
            double T3 = K1 * T1 * T2;
            return T3;
        }

        #endregion

        #region  弹性法算法
        private double Easy3(double K1, double K2,double K3)
        {

            double T1 = 1 + K2 * K3/100;
            double T2 = Math.Pow(T1, 5);
            double T3 = K1 * T2;
            return T3;

        }
        #endregion

        #region   天然气需求量算法
        private double Easy4(double K1, double K2)
        {
            double T1 = K1 * Math.Pow(10, 7) * 7000 / 8500 / Math.Pow(10, 8) * K2;
            return T1;
            
        }
            #endregion




        private void Calculate()
        {
            #region   基础数据输入的计算
            //城镇人口增长率计算   a[]是城镇人口数  E列
            string[] a = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                a[i] = (this.dataGridView1.Rows[i].Cells[2].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView1.RowCount - 2; i++)
            {
                this.dataGridView1.Rows[i + 1].Cells[3].Value = ((Convert.ToDouble(a[1 + i]) - Convert.ToDouble(a[i])) * 100 / Convert.ToDouble(a[i])).ToString("0.0");
            }
            //城镇化比率     b[]是总人口数  C列
            string[] b = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                b[i] = (this.dataGridView1.Rows[i].Cells[0].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                this.dataGridView1.Rows[i].Cells[4].Value = ((Convert.ToDouble(a[i]) / Convert.ToDouble(b[i])) * 100).ToString("0.0");
            }

            //人均GDP计算 c[]是 地区生产总值   H列
            string[] c = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                c[i] = (this.dataGridView1.Rows[i].Cells[5].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                this.dataGridView1.Rows[i].Cells[7].Value = ((Convert.ToDouble(c[i]) / Convert.ToDouble(b[i])) * 10000).ToString("0");
            }
            //能源消费增长 d[]是能源消费总量  K列
            string[] d = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                d[i] = (this.dataGridView1.Rows[i].Cells[8].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView1.RowCount - 2; i++)
            {
                this.dataGridView1.Rows[i + 1].Cells[13].Value = ((Convert.ToDouble(d[1 + i]) - Convert.ToDouble(d[i])) * 100 / Convert.ToDouble(d[i])).ToString("0.0");
            }

            //人均能源消费  e[]是人均能源消费  是R列
            string[] e = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                this.dataGridView1.Rows[i + 1].Cells[15].Value = (Convert.ToDouble(d[1 + i]) / Convert.ToDouble(b[i + 1])).ToString("0.00");
            }
            //电力消费增长  f[]是电力消费总量 T列
            string[] f = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                f[i] = (this.dataGridView1.Rows[i].Cells[17].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView1.RowCount - 2; i++)
            {
                this.dataGridView1.Rows[i + 1].Cells[18].Value = ((Convert.ToDouble(f[1 + i]) - Convert.ToDouble(f[i])) * 100 / Convert.ToDouble(f[i])).ToString("0.0");
            }

            //万元GDP电耗   g[]是  W列
            string[] g = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                this.dataGridView1.Rows[i + 1].Cells[20].Value = (Convert.ToDouble(f[1 + i]) / Convert.ToDouble(c[i + 1])).ToString("0.00");
            }
           
            //自然增长率 D列
            string[] h = new string[1000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                h[i] = (this.dataGridView1.Rows[i].Cells[1].Value).ToString();
            }
            this.dataGridView6.Rows[0].Cells[0].Value = (this.dataGridView1.Rows[5].Cells[2].Value).ToString();
            #endregion

            #region  参数设置的计算
            //参数设置页面的计算人口参数设定
            for (int i = 0; i < 3; i++)
            {
                this.dataGridView2.Rows[0].Cells[i].Value = ((Convert.ToDouble(this.dataGridView1.Rows[5].Cells[4].Value)) / 100).ToString("0.00");

            }
            //经济参数设定

            for (int i = 0; i < 3; i++)
            {
                this.dataGridView3.Rows[0].Cells[i].Value = ((Convert.ToDouble(this.dataGridView1.Rows[5].Cells[6].Value) - 100)).ToString("0.00");

            }

            //能源消费增长

            for (int i = 0; i < 3; i++)
            {
                this.dataGridView4.Rows[0].Cells[i].Value = ((Convert.ToDouble(this.dataGridView1.Rows[5].Cells[13].Value)).ToString("0.00"));
            }
            #endregion

            #region 预测结果的计算
            
            #region  城镇人口估算
            //D 列  高方案  
            //城镇人口计算
            string[] m = new string[1000];
            for (int i = 0; i < this.dataGridView2.RowCount - 1; i++)
            {
                m[i] = (this.dataGridView2.Rows[i].Cells[0].Value).ToString();
            }


            for (int i = 0; i < this.dataGridView6.RowCount - 2; i++)
            {
                this.dataGridView6.Rows[1 + i].Cells[0].Value = Easy(Convert.ToDouble(this.dataGridView6.Rows[i].HeaderCell.Value), Convert.ToDouble(m[i+1])).ToString("0.0");
            }
            //基础方案
            //城镇人口计算
            string[] n = new string[1000];
            for (int i = 0; i < this.dataGridView2.RowCount - 1; i++)
            {
                n[i] = (this.dataGridView2.Rows[i].Cells[1].Value).ToString();
            }


            for (int i = 0; i < this.dataGridView7.RowCount - 1; i++)
            {
                this.dataGridView7.Rows[i].Cells[0].Value = Easy(Convert.ToDouble(this.dataGridView7.Rows[i].HeaderCell.Value), Convert.ToDouble(n[i])).ToString("0.0");
            }

            //低方案
            //城镇人口计算
            string[] o = new string[1000];
            for (int i = 0; i < this.dataGridView2.RowCount - 1; i++)
            {
                o[i] = (this.dataGridView2.Rows[i].Cells[2].Value).ToString();
            }
            
            for (int i = 0; i < this.dataGridView8.RowCount - 1; i++)
            {
                this.dataGridView8.Rows[i].Cells[0].Value = Easy(Convert.ToDouble(this.dataGridView8.Rows[i].HeaderCell.Value), Convert.ToDouble(o[i])).ToString("0.0");
            }
            #endregion

            #region  增长法
            this.dataGridView6.Rows[0].Cells[1].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();
            // N 列 参数能源设置
            //高方案
            string[] k = new string[1000];
            for (int i = 0; i < this.dataGridView4.RowCount - 1; i++)
            {
                k[i] = (this.dataGridView4.Rows[i].Cells[0].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView6.RowCount - 2; i++)
            {
                this.dataGridView6.Rows[1 + i].Cells[1].Value = (Easy1(Convert.ToDouble(this.dataGridView6.Rows[i].Cells[1].Value), Convert.ToDouble(k[i]))).ToString("0.0");
            }
            //基础方案
            this.dataGridView7.Rows[0].Cells[1].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();

            string[] k1 = new string[1000];
            for (int i = 0; i < this.dataGridView4.RowCount - 1; i++)
            {
                k1[i] = (this.dataGridView4.Rows[i].Cells[1].Value).ToString();
            }


            for (int i = 0; i < this.dataGridView7.RowCount - 2; i++)
            {
                this.dataGridView7.Rows[1 + i].Cells[1].Value = (Easy1(Convert.ToDouble(this.dataGridView7.Rows[i].Cells[1].Value), Convert.ToDouble(k1[i]))).ToString("0.0");
            }
            //低方案

            this.dataGridView8.Rows[0].Cells[1].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();

            string[] k2 = new string[1000];
            for (int i = 0; i < this.dataGridView4.RowCount - 1; i++)
            {
                k2[i] = (this.dataGridView4.Rows[i].Cells[2].Value).ToString();
            }


            for (int i = 0; i < this.dataGridView8.RowCount - 2; i++)
            {
                this.dataGridView8.Rows[1 + i].Cells[1].Value = (Easy1(Convert.ToDouble(this.dataGridView8.Rows[i].Cells[1].Value), Convert.ToDouble(k2[i]))).ToString("0.0");
            }









            #endregion

            #region 能耗法
            //高方案
            this.dataGridView6.Rows[0].Cells[2].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();
            //经济参数  I列
            string[] k3 = new string[2000];
            for (int i = 0; i < this.dataGridView3.RowCount - 1; i++)
            {
                k3[i] = (this.dataGridView3.Rows[i].Cells[0].Value).ToString();
            }

            //能源参数 降耗指标 S列   
            string[] k4 = new string[2000];
            for (int i = 0; i < this.dataGridView5.RowCount - 1; i++)
            {
                k4[i] = (this.dataGridView5.Rows[i].Cells[0].Value).ToString();
            }
            for (int i = 0; i < this.dataGridView6.RowCount - 2; i++)
            {
                this.dataGridView6.Rows[1 + i].Cells[2].Value = (Easy2(Convert.ToDouble(this.dataGridView6.Rows[i].Cells[2].Value), Convert.ToDouble(k3[i]),Convert.ToDouble(k4[i]))).ToString("0.0");
            }



            //基础方案
            this.dataGridView7.Rows[0].Cells[2].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();
            //经济参数  j列
            string[] k5 = new string[2000];
            for (int i = 0; i < this.dataGridView3.RowCount - 1; i++)
            {
                k5[i] = (this.dataGridView3.Rows[i].Cells[1].Value).ToString();
            }

            //能源参数 降耗指标 t列   
            string[] k6 = new string[2000];
            for (int i = 0; i < this.dataGridView5.RowCount - 1; i++)
            {
                k6[i] = (this.dataGridView5.Rows[i].Cells[1].Value).ToString();
            }

            for (int i = 0; i < this.dataGridView7.RowCount - 2; i++)
            {
                this.dataGridView7.Rows[1 + i].Cells[2].Value = (Easy2(Convert.ToDouble(this.dataGridView7.Rows[i].Cells[2].Value), Convert.ToDouble(k5[i]), Convert.ToDouble(k6[i]))).ToString("0.0");
            }

            //低方案
            this.dataGridView8.Rows[0].Cells[2].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();
            //经济参数  K列
            string[] k7 = new string[2000];
            for (int i = 0; i < this.dataGridView3.RowCount - 1; i++)
            {
                k7[i] = (this.dataGridView3.Rows[i].Cells[2].Value).ToString();
            }

            //能源参数 降耗指标 U列   
            string[] k8 = new string[2000];
            for (int i = 0; i < this.dataGridView5.RowCount - 1; i++)
            {
                k8[i] = (this.dataGridView5.Rows[i].Cells[2].Value).ToString();
            }

            for (int i = 0; i < this.dataGridView8.RowCount - 2; i++)
            {
                this.dataGridView8.Rows[1 + i].Cells[2].Value = (Easy2(Convert.ToDouble(this.dataGridView8.Rows[i].Cells[2].Value), Convert.ToDouble(k7[i]), Convert.ToDouble(k8[i]))).ToString("0.0");
            }





            #endregion

            #region  弹性法
            //高方案
            //Q 列   能源弹性
            string[] k10 = new string[2000];
            for (int i = 0; i < this.dataGridView1.RowCount - 1; i++)
            {
                k10[i] = (this.dataGridView1.Rows[i].Cells[14].Value).ToString();
            }
            
            this.dataGridView6.Rows[0].Cells[3].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();
            double S1 = 0;

            int T1 = Convert.ToInt32(textBox2.Text);

            for (int i = 0; i < 5; i++)
            {
                S1 = S1 +Convert.ToDouble( k10[T1 - 2010 - i]);
            }
            double Var = S1 / 5;
            for (int i = 0; i < this.dataGridView6.RowCount - 2; i++)
            {
                this.dataGridView6.Rows[1 + i].Cells[3].Value = (Easy3(Convert.ToDouble(this.dataGridView6.Rows[i].Cells[3].Value), Var, Convert.ToDouble(k3[i]))).ToString("0.0");
            }
            //  基础方案、
            this.dataGridView7.Rows[0].Cells[3].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();
            
            for (int i = 0; i < this.dataGridView7.RowCount - 2; i++)
            {
                this.dataGridView7.Rows[1 + i].Cells[3].Value = (Easy3(Convert.ToDouble(this.dataGridView7.Rows[i].Cells[3].Value), Var, Convert.ToDouble(k5[i]))).ToString("0.0");
            }
            //低方案
            this.dataGridView8.Rows[0].Cells[3].Value = (this.dataGridView1.Rows[5].Cells[8].Value).ToString();

            for (int i = 0; i < this.dataGridView8.RowCount - 2; i++)
            {
                this.dataGridView8.Rows[1 + i].Cells[3].Value = (Easy3(Convert.ToDouble(this.dataGridView8.Rows[i].Cells[3].Value), Var, Convert.ToDouble(k7[i]))).ToString("0.0");
            }
            #endregion

            #region  平均值计算

            for (int i = 0; i < this.dataGridView6.RowCount-1; i++)
            {
                this.dataGridView6.Rows[i].Cells[4].Value = ((Convert.ToDouble(this.dataGridView6.Rows[i].Cells[1].Value) + Convert.ToDouble(this.dataGridView6.Rows[i].Cells[2].Value) + Convert.ToDouble(this.dataGridView6.Rows[i].Cells[3].Value)) / 3).ToString("0.0");

                this.dataGridView7.Rows[i].Cells[4].Value = ((Convert.ToDouble(this.dataGridView7.Rows[i].Cells[1].Value) + Convert.ToDouble(this.dataGridView7.Rows[i].Cells[2].Value) + Convert.ToDouble(this.dataGridView7.Rows[i].Cells[3].Value)) / 3).ToString("0.0");
                this.dataGridView8.Rows[i].Cells[4].Value = ((Convert.ToDouble(this.dataGridView8.Rows[i].Cells[1].Value) + Convert.ToDouble(this.dataGridView8.Rows[i].Cells[2].Value) + Convert.ToDouble(this.dataGridView8.Rows[i].Cells[3].Value)) / 3).ToString("0.0");


            }









            #endregion

            #endregion
        }







        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "增长法")
            {
                for (int i = 0; i < this.dataGridView6.RowCount - 1; i++)
                {
                    this.dataGridView6.Rows[i].Cells[5].Value = (this.dataGridView6.Rows[i].Cells[1].Value);
                    this.dataGridView7.Rows[i].Cells[5].Value = (this.dataGridView7.Rows[i].Cells[1].Value);
                    this.dataGridView8.Rows[i].Cells[5].Value = (this.dataGridView8.Rows[i].Cells[1].Value);
                }
                
            }
            if (comboBox2.Text == "能耗法")
            {
                for (int i = 0; i < this.dataGridView6.RowCount - 1; i++)
                {
                    
                    this.dataGridView6.Rows[i].Cells[5].Value = (this.dataGridView6.Rows[i].Cells[2].Value);
                    this.dataGridView7.Rows[i].Cells[5].Value = (this.dataGridView7.Rows[i].Cells[2].Value);
                    this.dataGridView8.Rows[i].Cells[5].Value = (this.dataGridView8.Rows[i].Cells[2].Value);
                }
            }

            if (comboBox2.Text == "弹性法")
            {
                for (int i = 0; i < this.dataGridView6.RowCount - 1; i++)
                {

                    this.dataGridView6.Rows[i].Cells[5].Value = (this.dataGridView6.Rows[i].Cells[3].Value);
                    this.dataGridView7.Rows[i].Cells[5].Value = (this.dataGridView7.Rows[i].Cells[3].Value);
                    this.dataGridView8.Rows[i].Cells[5].Value = (this.dataGridView8.Rows[i].Cells[3].Value);
                }
            }

            if (comboBox2.Text == "平均值")
            {
                for (int i = 0; i < this.dataGridView6.RowCount - 1; i++)
                {
                    this.dataGridView6.Rows[i].Cells[5].Value = (this.dataGridView6.Rows[i].Cells[4].Value);
                    this.dataGridView7.Rows[i].Cells[5].Value = (this.dataGridView7.Rows[i].Cells[4].Value);
                    this.dataGridView8.Rows[i].Cells[5].Value = (this.dataGridView8.Rows[i].Cells[4].Value);
                }
            }

            //高方案 天然气需求
            string[] n1 = new string[1000];
            for (int i = 0; i < this.dataGridView5.RowCount - 1; i++)
            {
                n1[i] = (this.dataGridView5.Rows[i].Cells[3].Value).ToString();
            }
            
            for (int i = 0; i < this.dataGridView6.RowCount - 1; i++)
            {
                this.dataGridView6.Rows[i].Cells[6].Value = (Easy4(Convert.ToDouble(this.dataGridView6.Rows[i].Cells[5].Value),Convert.ToDouble(n1[i]))).ToString("0.0");
            }
            //基础方案  天然气需求
            for (int i = 0; i < this.dataGridView7.RowCount - 1; i++)
            {
                this.dataGridView7.Rows[i].Cells[6].Value = (Easy4(Convert.ToDouble(this.dataGridView7.Rows[i].Cells[5].Value), Convert.ToDouble(n1[i]))).ToString("0.0");
            }
            //低方案
            for (int i = 0; i < this.dataGridView8.RowCount - 1; i++)
            {
                this.dataGridView8.Rows[i].Cells[6].Value = (Easy4(Convert.ToDouble(this.dataGridView8.Rows[i].Cells[5].Value), Convert.ToDouble(n1[i]))).ToString("0.0");
            }
            //百分差的计算

            for (int i = 0; i < this.dataGridView6.RowCount - 1; i++)
            {
                this.dataGridView6.Rows[i].Cells[8].Value =(((Convert.ToDouble(this.dataGridView6.Rows[i].Cells[5].Value))-Convert.ToDouble(this.dataGridView6.Rows[i].Cells[6].Value))/ Convert.ToDouble(this.dataGridView6.Rows[i].Cells[5].Value)).ToString("0.0");
                
            }
            for (int i = 0; i < this.dataGridView7.RowCount - 1; i++)
            {
                this.dataGridView7.Rows[i].Cells[8].Value = (((Convert.ToDouble(this.dataGridView7.Rows[i].Cells[5].Value)) - Convert.ToDouble(this.dataGridView7.Rows[i].Cells[6].Value)) / Convert.ToDouble(this.dataGridView7.Rows[i].Cells[5].Value)).ToString("0.0");

            }


            for (int i = 0; i < this.dataGridView8.RowCount - 1; i++)
            {
                this.dataGridView8.Rows[i].Cells[8].Value = (((Convert.ToDouble(this.dataGridView8.Rows[i].Cells[5].Value)) - Convert.ToDouble(this.dataGridView8.Rows[i].Cells[6].Value)) / Convert.ToDouble(this.dataGridView8.Rows[i].Cells[5].Value)).ToString("0.0");

            }





        }
    }
    #endregion
}

