using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 天然气市场需求分析软件_求你不死机版_
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }
        public Windows7 EnergyConsumptionAnalys;
        public Windows1 CityGas;
        public Windows3 GenerateElectricity;
        public Windows5 Traffic;
        public Windows10 SmartAnalysCityGas;
        public Windows6 SavingEnergy;
        public Windows4 Chemical;
        public Windows9 KilnConsumptionGas;
        public DistributeEnergy1 Distributedpoweranalysis;
        public SupportingDocumentJudgmentModelRegionalStyle SupportingdocumentjudgmentModelRegionalStyle;
        public SupportingDocumentJudgmentModelFloorStyle SupportingdocumentJudgmentModelFloorStyle;
        public RiskJudgmentModelRegionalStyle RiskJudgmentmodelRegionalStyle;
        public PolicyCountry Policycountry;
        public PolicyRural Policyrural;
        private string skinPath = @"Resources\";
        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 关闭所有ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 交通ToolStripMenuItem_Click(object sender, EventArgs e)
        {  
            Traffic = new Windows5();
            Traffic.MdiParent = this;
            Traffic.Show();
        }

        private void 能耗系数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EnergyConsumptionAnalys = new Windows7();
            EnergyConsumptionAnalys.MdiParent = this;
            EnergyConsumptionAnalys.Show();
        }

        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SendKeys.Send("^{C}");

        }

        private void 城市燃气ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SmartAnalysCityGas = new Windows10();
            SmartAnalysCityGas.MdiParent = this;
            SmartAnalysCityGas.Show();
        }
        
        private void 压气站布置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GenerateElectricity = new Windows3();
            GenerateElectricity.MdiParent = this;
            GenerateElectricity.Show();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.toolStripStatusLabel3.Text = "系统当前时间：" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            SetImg();   //为菜单子选项设置图标

            //this.WindowState == FormWindowState.Maximized;
            this.toolStripStatusLabel3.Text = "系统当前时间：" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

            this.timer1.Interval = 1000;

            this.timer1.Start();
        }
        private void SetImg()
        {
            //文件菜单的图标设置
            新建ToolStripMenuItem.Image = Image.FromFile(skinPath + "New.png");
            打开ToolStripMenuItem.Image = Image.FromFile(skinPath + "Open.png");
            关闭ToolStripMenuItem.Image = Image.FromFile(skinPath + "Close.png");
            关闭所有ToolStripMenuItem.Image = Image.FromFile(skinPath + "CloseAll.png");
            保存ToolStripMenuItem.Image = Image.FromFile(skinPath + "Save.png");
            另存为ToolStripMenuItem.Image = Image.FromFile(skinPath + "SaveAs.png");
            最近文件ToolStripMenuItem.Image = Image.FromFile(skinPath + "RecentDocuments.png");
            退出ToolStripMenuItem.Image = Image.FromFile(skinPath + "Exit.png");

            toolStripLabel1.Image = Image.FromFile(skinPath + "New.png");
            toolStripLabel3.Image = Image.FromFile(skinPath + "Open.png");
            toolStripLabel2.Image = Image.FromFile(skinPath + "Save.png");
            toolStripLabel7.Image = Image.FromFile(skinPath + "Cut.png");
            toolStripLabel6.Image = Image.FromFile(skinPath + "Copy.png");
            toolStripLabel5.Image = Image.FromFile(skinPath + "Paste.png");
            toolStripLabel4.Image = Image.FromFile(skinPath + "Search.png");
            toolStripLabel8.Image = Image.FromFile(skinPath + "Help.png");

            foreach (ToolStripItem vItem in toolStrip1.Items)
            {
                if (vItem is ToolStripLabel)
                {
                    vItem.Text = "";
                    vItem.AutoSize = false;
                    vItem.Width = 40;
                }
            }

        }


        private void 智能分析城市燃气ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SmartAnalysCityGas = new Windows10();
            SmartAnalysCityGas.MdiParent = this;
            SmartAnalysCityGas.Show();
        }

        private void 节能减排ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void 状态栏ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem12.Checked = !ToolStripMenuItem12.Checked;
            statusStrip1.Visible = !statusStrip1.Visible;
        }

        private void ToolStripMenuItem22_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem22.Checked = !ToolStripMenuItem22.Checked;
            statusStrip1.Visible = !statusStrip1.Visible;
        }

        private void 化工ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Chemical = new Windows4();
            Chemical.MdiParent = this;
            Chemical.Show();
        }

        private void 锅炉用气ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //BoilerConsumptionGas = new Windows8();
            //BoilerConsumptionGas.MdiParent = this;
            //BoilerConsumptionGas.Show();
        }

        private void 窑炉用气toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            //KilnConsumptionGas = new Windows9();
            //KilnConsumptionGas.MdiParent = this;
            //KilnConsumptionGas.Show();
        }

        private void 层叠窗口ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.Cascade);//层叠
        }

        private void 水平平铺ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileHorizontal); //水平
        }

        private void 垂直平铺ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileVertical);  //垂直
        }

        private void 工业燃料ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KilnConsumptionGas = new Windows9();
            KilnConsumptionGas.MdiParent = this;
            KilnConsumptionGas.Show();
        }
        private void 管道分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 技术经济型模型ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Distributedpoweranalysis = new DistributeEnergy1();
            Distributedpoweranalysis.MdiParent = this;
            Distributedpoweranalysis.Show();
        }

        private void 区域式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SupportingdocumentjudgmentModelRegionalStyle = new SupportingDocumentJudgmentModelRegionalStyle();
            SupportingdocumentjudgmentModelRegionalStyle.MdiParent = this;
            SupportingdocumentjudgmentModelRegionalStyle.Show();
        }

        private void 楼宇式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SupportingdocumentJudgmentModelFloorStyle = new SupportingDocumentJudgmentModelFloorStyle();
            SupportingdocumentJudgmentModelFloorStyle.MdiParent = this;
            SupportingdocumentJudgmentModelFloorStyle.Show();
        }

        private void 区域式ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RiskJudgmentmodelRegionalStyle = new RiskJudgmentModelRegionalStyle();
            RiskJudgmentmodelRegionalStyle.MdiParent = this;
            RiskJudgmentmodelRegionalStyle.Show();
        }

        private void 国家层面ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Policycountry = new PolicyCountry();
            Policycountry.MdiParent = this;
            Policycountry.Show();
        }

        private void 地方层面ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            Policyrural = new PolicyRural();
            Policyrural.MdiParent = this;
            Policyrural.Show();
        }

        private void 节能减排ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SavingEnergy = new Windows6();
            SavingEnergy.MdiParent = this;
            SavingEnergy.Show();
        }
    }
}
