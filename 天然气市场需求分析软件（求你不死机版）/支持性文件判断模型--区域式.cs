using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace 天然气市场需求分析软件_求你不死机版_
{

    public partial class SupportingDocumentJudgmentModelRegionalStyle : Form
    {
        public SupportingDocumentJudgmentModelRegionalStyle()
        {
            InitializeComponent();
        }

        private void SupportingDocumentJudgmentModelRegionalStyle_Load(object sender, EventArgs e)
        {
            int index = this.dataGridView1.Rows.Add(15);
            this.dataGridView1.Rows[0].Cells[0].Value = "序号";
            this.dataGridView1.Rows[1].Cells[0].Value = "序号";
            this.dataGridView1.Rows[0].Cells[1].Value = "立项判断条件";
            this.dataGridView1.Rows[1].Cells[1].Value = "立项判断条件";
            this.dataGridView1.Rows[0].Cells[2].Value = "文件取得方或签订方";
            this.dataGridView1.Rows[1].Cells[2].Value = "文件取得方或签订方";


            this.dataGridView1.Rows[2].Cells[0].Value = "一";
            this.dataGridView1.Rows[3].Cells[0].Value = "1";
            this.dataGridView1.Rows[4].Cells[0].Value = "2";
            this.dataGridView1.Rows[5].Cells[0].Value = "3";
            this.dataGridView1.Rows[6].Cells[0].Value = "4";
            this.dataGridView1.Rows[7].Cells[0].Value = "5";
            this.dataGridView1.Rows[8].Cells[0].Value = "6";
            this.dataGridView1.Rows[9].Cells[0].Value = "7";
            this.dataGridView1.Rows[10].Cells[0].Value = "二";
            this.dataGridView1.Rows[11].Cells[0].Value = "1";
            this.dataGridView1.Rows[12].Cells[0].Value = "2";
            this.dataGridView1.Rows[13].Cells[0].Value = "3";
            this.dataGridView1.Rows[14].Cells[0].Value = "4";

            this.dataGridView1.Rows[2].Cells[1].Value = "政府、集团支持性文件";
            this.dataGridView1.Rows[3].Cells[1].Value = "主体工程可研报告及相关专题报告编制完成，并通过审查";
            this.dataGridView1.Rows[4].Cells[1].Value = "应取得县级及以上主管部门同意开展前期工作的文件";
            this.dataGridView1.Rows[5].Cells[1].Value = "应取得县级及以上主管部门关于规划选址、用地预审的意见";
            this.dataGridView1.Rows[6].Cells[1].Value = "接入系统、环境保护、水土保持、水资源论证等专题报告编制完成，并且与主管部门沟通无颠覆性意见";
            this.dataGridView1.Rows[7].Cells[1].Value = "项目符合区域热电联产规划、供热规划要求";
            this.dataGridView1.Rows[8].Cells[1].Value = "应取得主管部门对军事设施不影响、对航空不影响、不压矿、不压文物等文件";
            this.dataGridView1.Rows[9].Cells[1].Value = "取得集团公司发起备案的批复文件";
            this.dataGridView1.Rows[10].Cells[1].Value = "相关协议";
            this.dataGridView1.Rows[11].Cells[1].Value = "签订天然气供应意向协议";
            this.dataGridView1.Rows[12].Cells[1].Value = "签订供热（冷）意向协议";
            this.dataGridView1.Rows[13].Cells[1].Value = "签订供电意向协议（如有)";
            this.dataGridView1.Rows[14].Cells[1].Value = "签订供水意向协议";


            this.dataGridView1.Rows[4].Cells[2].Value = "县级及以其上主管部门";
            this.dataGridView1.Rows[5].Cells[2].Value = "县级及以上主管部门";
            this.dataGridView1.Rows[8].Cells[2].Value = "相应级别的政府主管部门";
            this.dataGridView1.Rows[9].Cells[2].Value = "集团公司";

            this.dataGridView1.Rows[11].Cells[2].Value = "天然气供应方";
            this.dataGridView1.Rows[12].Cells[2].Value = "热（冷）用户";
            this.dataGridView1.Rows[13].Cells[2].Value = "电用户";
            this.dataGridView1.Rows[14].Cells[2].Value = "供水方";

            this.dataGridView1.Rows[0].Cells[3].Value = " 判断依据";
            this.dataGridView1.Rows[0].Cells[4].Value = " 判断依据";
            this.dataGridView1.Rows[1].Cells[3].Value = "在立项阶段落实";
            this.dataGridView1.Rows[1].Cells[4].Value = "可在下阶段落实";

        }
        #region //实现对于单元格的合并功能
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // 对第1列相同单元格进行合并
            //    if (e.ColumnIndex == 0 && e.RowIndex != -1|| e.ColumnIndex == 1 && e.RowIndex != -1|| e.ColumnIndex == 2 && e.RowIndex != -1)
            //    {
            //        using
            //            (
            //            Brush datagridBrush = new SolidBrush(dataGridView1.GridColor),
            //            backColorBrush = new SolidBrush(e.CellStyle.BackColor)
            //            )
            //        {
            //            using (Pen gridLinePen = new Pen(datagridBrush))
            //            {
            //                // 清除单元格
            //                //e.Graphics.FillRectangle(backColorBrush, e.CellBounds);

            //                //// 画 Grid 边线（仅画单元格的底边线和右边线）
            //                ////   如果下一行和当前行的数据不同，则在当前的单元格画一条底边线
            //                ////if (e.RowIndex < dataGridView1.Rows.Count - 1 &&
            //                ////dataGridView1.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value.ToString()!= e.Value.ToString())
            //                ////    e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left,
            //                ////    e.CellBounds.Bottom - 1, e.CellBounds.Right - 1,
            //                ////    e.CellBounds.Bottom - 1);
            //                // 画右边线
            //                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1,
            //                    e.CellBounds.Top, e.CellBounds.Right - 1,
            //                    e.CellBounds.Bottom);

            //                // 画（填写）单元格内容，相同的内容的单元格只填写第一个
            //                if (e.Value != null)
            //                {
            //                    if (e.RowIndex > 0 && dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() == e.Value.ToString())
            //                    {
            //                    }
            //                    else
            //                    {
            //                        e.Graphics.DrawString((String)e.Value, e.CellStyle.Font,
            //                        Brushes.Black, e.CellBounds.X + 2,
            //                        e.CellBounds.Y + 5, StringFormat.GenericDefault);
            //                    }
            //                }
            //                e.Handled = true;
            //            }
            //        }

            //    }


            //}
            #endregion
            //}
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}