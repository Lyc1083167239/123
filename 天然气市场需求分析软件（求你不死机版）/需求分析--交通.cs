using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;

namespace 天然气市场需求分析软件_求你不死机版_
{

    public partial class Windows5 : Form
    {


        string[,] cell = new string[25, 25];
        //string[] sheet1 = new string[10];
        double[] BL1 = new double[26];
        double[] BL2 = new double[26];
        double[] BL3 = new double[26];
        double[] BL4 = new double[26];
        double[] BL5 = new double[26];
        double[] BL6 = new double[26];
        double[] BL7 = new double[26];

        string filename = "";
        public Windows5()
        {
            InitializeComponent();
        }
        //private int[] GetRow(string s, int n)
        //{

        //}
        #region 
        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    //try {
            //    //    HSSFWorkbook workbook2003 = new HSSFWorkbook(); //新建工作簿
            //    //    workbook2003.CreateSheet("Sheet1");  //新建1个Sheet工作表            
            //    //    HSSFSheet SheetOne = (HSSFSheet)workbook2003.GetSheet("Sheet1"); //获取名称为Sheet1的工作表
            //    //    //对工作表先添加行，下标从0开始
            //    //    for (int i = 0; i < 10; i++)
            //    //    {
            //    //        SheetOne.CreateRow(i);   //创建10行
            //    //    }
            //    //    //对每一行创建10个单元格
            //    //    HSSFRow SheetRow = (HSSFRow)SheetOne.GetRow(0);  //获取Sheet1工作表的首行
            //    //    HSSFCell[] SheetCell = new HSSFCell[10];
            //    //    for (int i = 0; i < 10; i++)
            //    //    {
            //    //        SheetCell[i] = (HSSFCell)SheetRow.CreateCell(i);  //为第一行创建10个单元格
            //    //    }
            //    //    //创建之后就可以赋值了
            //    //    SheetCell[0].SetCellValue(true); //赋值为bool型         
            //    //    SheetCell[1].SetCellValue(0.000001); //赋值为浮点型
            //    //    SheetCell[2].SetCellValue("Excel2003"); //赋值为字符串
            //    //    SheetCell[3].SetCellValue("123456789987654321");//赋值为长字符串
            //    //    for (int i = 4; i < 10; i++)
            //    //    {
            //    //        SheetCell[i].SetCellValue(i);    //循环赋值为整形
            //    //    }
            //    //    SheetOne.GetRow(0).GetCell(1).SetCellValue("131");
            //    //    FileStream file2003 = new FileStream(@"E:\Excel2003.xls", FileMode.Create);
            //    //    workbook2003.Write(file2003);
            //    //    file2003.Close();
            //    //    workbook2003.Close();
            //    IWorkbook workbook = null;  //新建IWorkbook对象
            //    string fileName = filename;
            //    FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            //    if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            //    {
            //        workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            //    }
            //    else if (fileName.IndexOf(".xls") > 0) // 2003版本
            //    {
            //        workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
            //    }
            //    ISheet sheet = workbook.GetSheetAt(3);
            //    //获取第一个工作表                                                  
            //    //IRow row;// = sheet.GetRow(0);            
            //    //新建当前工作表行数据                                                   
            //    //for (int i = 0; i < sheet.LastRowNum; i++)  //对工作表每一行                                                 
            //    //{                                                 
            //    //    row = sheet.GetRow(i);   //row读入第i行数据                                                   
            //    //    if (row != null)                                                    
            //    //    {                                                   
            //    //        for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列                                                 
            //    //        {                                                  
            //    //            string cellValue = row.GetCell(j).ToString(); //获取i行j列数据                                                  
            //    //           Debug.WriteLine(cellValue);                                                   
            //    //        }                                                  
            //    //    }                                         
            //    //}
            //    #region 设置人口增长率
            //    double Increase = Convert.ToDouble(txtInput1.Text);

            //    #endregion
            //    #region   读取输入人口数据
            //    for (int s = 0; s < 11; s++)
            //    {
            //        for (int i = 0; i < 5; i++)
            //        {
            //            cell[s, i] = sheet.GetRow(s + 4).GetCell(i).ToString();
            //            //Debug.WriteLine(cell[i]);
            //        }
            //        for (int i = 5; i < 23; i++)
            //        {
            //            cell[s, i] = (Convert.ToDouble(cell[s, i-1]) * Increase).ToString("0.0000");
            //            //Debug.WriteLine(cell[i]);
            //        }
            //    }

            //            #endregion
            //            //Debug.WriteLine(cell[i]);
            //            //string cellValue1 = sheet.GetRow(4).GetCell(0).ToString(); //获取i行j列数据

            //            //Debug.WriteLine(cellValue1);
            //            Console.ReadLine();
            //    fileStream.Close();
            //    workbook.Close();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            //IWorkbook workbook1 = null;  //新建IWorkbook对象
            //string fileProcess = "智能分析-城市燃气.xlsx";
            //FileStream fileProcessStream1 = new FileStream(fileProcess, FileMode.Open, FileAccess.Read);
            //if (fileProcess.IndexOf(".xlsx") > 0) // 2007版本
            //{
            //    workbook1 = new XSSFWorkbook(fileProcessStream1);  //xlsx数据读入workbook
            //    fileProcessStream1.Close();
            //}
            //else if (fileProcess.IndexOf(".xls") > 0) // 2003版本
            //{
            //    workbook1 = new HSSFWorkbook(fileProcessStream1);  //xls数据读入workbook

            //}
            ////string[] sheets = new string[10];
            //ISheet[] sheet1 = new ISheet[12];
            //for (int a = 0; a < 11; a++)
            //{
            //    sheet1[a] = workbook1.GetSheetAt(a);  //获取第一个工作表  


            //    sheet1[a].GetRow(1).GetCell(0).SetCellValue(cell[a, 2]);

            //    for (int k = 4; k < 25; k++)
            //    {
            //        BL1[k] = Convert.ToDouble(cell[a, k - 1]);//15行数据
            //        BL2[k] = Convert.ToDouble(sheet1[a].GetRow(15).GetCell(k).ToString());//16行数据
            //                                                                              //double BL3 = Convert.ToDouble(sheet1.GetRow(19).GetCell(k));
            //        sheet1[a].GetRow(16).GetCell(k).SetCellValue(BL1[k] * BL2[k]);//填充17行数据
            //        sheet1[a].GetRow(19).GetCell(k).SetCellValue(BL1[k] * 30);//填充20行数据 
            //                                                                  //BL4[k] = Convert.ToDouble(sheet1.GetRow(19).GetCell(k).ToString ());//20行数据
            //        BL3[k] = Convert.ToDouble(sheet1[a].GetRow(20).GetCell(k).ToString());
            //        sheet1[a].GetRow(21).GetCell(k).SetCellValue(BL1[k] * 30 * BL3[k]);//填充22行数据
            //        BL4[k] = Convert.ToDouble(sheet1[a].GetRow(17).GetCell(k).ToString());//18行数据
            //        BL5[k] = Convert.ToDouble(sheet1[a].GetRow(18).GetCell(k).ToString());//18行数据
            //        sheet1[a].GetRow(2).GetCell(k).SetCellValue(BL1[k] * BL2[k] * 60 / 10000);//填充3行数据
            //        sheet1[a].GetRow(3).GetCell(k).SetCellValue(BL1[k] * BL2[k] * 60 / 10000 * BL4[k]);//填充4行数据
            //        sheet1[a].GetRow(4).GetCell(k).SetCellValue(BL1[k] * BL3[k] * 30 * 12 / 10000);//填充5行数据
            //        sheet1[a].GetRow(5).GetCell(k).SetCellValue(BL1[k] * BL2[k] * 60 / 10000 * BL5[k]);//填充6行数据
            //        BL7[k] = BL1[k] * BL2[k] * 60 / 10000 + BL1[k] * BL2[k] * 60 / 10000 * BL4[k] + BL1[k] * BL3[k] * 30 * 12 / 10000 + BL1[k] * BL2[k] * 60 / 10000 * BL5[k];
            //        sheet1[a].GetRow(6).GetCell(k).SetCellValue(BL7[k]);//填7行数据
            //    }
            //    for (int j = 1; j < 22; j++)
            //    {
            //        sheet1[a].GetRow(14).GetCell(j + 1).SetCellValue(cell[a, j]);
            //    }
            //}

            //FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\计算模板.xlsx", FileMode.OpenOrCreate);
            //workbook1.Write(file);
            //file.Close();
            //workbook1.Close();
        }
        #endregion

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtOutput1.Text = openFileDialog1.FileName;
                filename = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                //try {
                //    HSSFWorkbook workbook2003 = new HSSFWorkbook(); //新建工作簿
                //    workbook2003.CreateSheet("Sheet1");  //新建1个Sheet工作表            
                //    HSSFSheet SheetOne = (HSSFSheet)workbook2003.GetSheet("Sheet1"); //获取名称为Sheet1的工作表
                //    //对工作表先添加行，下标从0开始
                //    for (int i = 0; i < 10; i++)
                //    {
                //        SheetOne.CreateRow(i);   //创建10行
                //    }
                //    //对每一行创建10个单元格
                //    HSSFRow SheetRow = (HSSFRow)SheetOne.GetRow(0);  //获取Sheet1工作表的首行
                //    HSSFCell[] SheetCell = new HSSFCell[10];
                //    for (int i = 0; i < 10; i++)
                //    {
                //        SheetCell[i] = (HSSFCell)SheetRow.CreateCell(i);  //为第一行创建10个单元格
                //    }
                //    //创建之后就可以赋值了
                //    SheetCell[0].SetCellValue(true); //赋值为bool型         
                //    SheetCell[1].SetCellValue(0.000001); //赋值为浮点型
                //    SheetCell[2].SetCellValue("Excel2003"); //赋值为字符串
                //    SheetCell[3].SetCellValue("123456789987654321");//赋值为长字符串
                //    for (int i = 4; i < 10; i++)
                //    {
                //        SheetCell[i].SetCellValue(i);    //循环赋值为整形
                //    }
                //    SheetOne.GetRow(0).GetCell(1).SetCellValue("131");
                //    FileStream file2003 = new FileStream(@"E:\Excel2003.xls", FileMode.Create);
                //    workbook2003.Write(file2003);
                //    file2003.Close();
                //    workbook2003.Close();
                IWorkbook workbook = null;  //新建IWorkbook对象
                //if (openFileDialog1.ShowDialog() == DialogResult.OK)
                //{
                //    txtOutput1.Text = openFileDialog1.FileName;
                //    filename = openFileDialog1.FileName;
                //}
                string fileName = txtOutput1.Text;
                FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
                }
                ISheet sheet = workbook.GetSheetAt(3);
                //获取第一个工作表                                                  
                //IRow row;// = sheet.GetRow(0);            
                //新建当前工作表行数据                                                   
                //for (int i = 0; i < sheet.LastRowNum; i++)  //对工作表每一行                                                 
                //{                                                 
                //    row = sheet.GetRow(i);   //row读入第i行数据                                                   
                //    if (row != null)                                                    
                //    {                                                   
                //        for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列                                                 
                //        {                                                  
                //            string cellValue = row.GetCell(j).ToString(); //获取i行j列数据                                                  
                //           Debug.WriteLine(cellValue);                                                   
                //        }                                                  
                //    }                                         
                //}
                #region 设置人口增长率
                double Increase = Convert.ToDouble(txtInput1.Text);
                #endregion
                #region   读取输入人口数据
                for (int s = 0; s < 11; s++)
                {
                    for (int i = 0; i < 3; i++)
                    {
                        cell[s, i] = sheet.GetRow(s + 4).GetCell(i).ToString();
                        //Debug.WriteLine(cell[i]);
                    }
                    for (int i = 3; i < 22; i++)
                    {
                        cell[s, i] = (Convert.ToDouble(cell[s, i - 1]) * Increase).ToString("0.0000");
                        //Debug.WriteLine(cell[i]);
                    }
                }
                #endregion
                //Debug.WriteLine(cell[i]);
                //string cellValue1 = sheet.GetRow(4).GetCell(0).ToString(); //获取i行j列数据

                //Debug.WriteLine(cellValue1);
                Console.ReadLine();
                fileStream.Close();
                workbook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            IWorkbook workbook1 = null;  //新建IWorkbook对象
            string fileProcess = "需求预测交通1.xlsx";
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
            //string[] sheets = new string[10];
            ISheet[] sheet1 = new ISheet[12];
            for (int a = 0; a < 11; a++)
            {
                sheet1[a] = workbook1.GetSheetAt(a);  //获取第一个工作表  


                sheet1[a].GetRow(0).GetCell(23).SetCellValue(cell[a, 0]);

                for (int k = 2; k < 23; k++)
                {
                    BL1[k] = Convert.ToDouble(cell[a, k - 1]);//15行数据
                    BL2[k] = Convert.ToDouble(sheet1[a].GetRow(15).GetCell(k).ToString());//16行数据
                                                                                          //double BL3 = Convert.ToDouble(sheet1.GetRow(19).GetCell(k));
                    sheet1[a].GetRow(16).GetCell(k).SetCellValue(BL1[k] * BL2[k]);//填充17行数据
                    sheet1[a].GetRow(19).GetCell(k).SetCellValue(BL1[k] * 30);//填充20行数据 
                                                                              //BL4[k] = Convert.ToDouble(sheet1.GetRow(19).GetCell(k).ToString ());//20行数据
                    BL3[k] = Convert.ToDouble(sheet1[a].GetRow(20).GetCell(k).ToString());
                    sheet1[a].GetRow(21).GetCell(k).SetCellValue(BL1[k] * 30 * BL3[k]);//填充22行数据
                    BL4[k] = Convert.ToDouble(sheet1[a].GetRow(17).GetCell(k).ToString());//18行数据
                    BL5[k] = Convert.ToDouble(sheet1[a].GetRow(18).GetCell(k).ToString());//18行数据
                    //sheet1[a].GetRow(2).GetCell(k).SetCellValue(BL1[k] * BL2[k] * 60 / 10000);//填充3行数据
                    //sheet1[a].GetRow(3).GetCell(k).SetCellValue(BL1[k] * BL2[k] * 60 / 10000 * BL4[k]);//填充4行数据
                    //sheet1[a].GetRow(4).GetCell(k).SetCellValue(BL1[k] * BL3[k] * 30 * 12 / 10000);//填充5行数据
                    sheet1[a].GetRow(2).GetCell(k).SetCellValue(BL1[k] * BL2[k] * 60 / 10000 * BL5[k]);//填充6行数据
                    BL7[k] = BL1[k] * BL2[k] * 60 / 10000 + BL1[k] * BL2[k] * 60 / 10000 * BL4[k] + BL1[k] * BL3[k] * 30 * 12 / 10000 + BL1[k] * BL2[k] * 60 / 10000 * BL5[k];
                    //sheet1[a].GetRow(6).GetCell(k).SetCellValue(BL7[k]);//填7行数据
                }
                for (int j = 1; j < 22; j++)
                {
                    sheet1[a].GetRow(14).GetCell(j + 1).SetCellValue(cell[a, j]);
                }
            }

            FileStream file = new FileStream("C:\\Users\\Administrator\\Desktop\\需求预测-交通333.xlsx", FileMode.OpenOrCreate);
            workbook1.Write(file);
            file.Close();
            workbook1.Close();
        }

        private void txtOutput1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Windows1_Load(object sender, EventArgs e)
        {

        }
        
    }
}
