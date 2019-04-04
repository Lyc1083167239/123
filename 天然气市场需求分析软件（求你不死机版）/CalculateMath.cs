using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 天然气市场需求分析软件_求你不死机版_
{
    class CalculateMath
    {
        #region     总收入1+2+3
        //txtInput10   txtInput11   txtInput7    txtInput16    txtInput8    txtInput9 txtInput18
        public double AllSum(double T1, double T2, double T3, double T4, double T5, double T7, double T8, double T9)
        {
            //Double P8 = Convert.ToDouble();//年利用小时数
            //Double P9 = Convert.ToDouble(.Text) / 100;//负荷率
            //txtOutput12
            Double K1 = T3 * T1 / 10000.0 * T2 / 100.0; //发电量
            //txtOutput11
            Double K2 = T4 * K1 / 1.0;//1-电费
                                      //Double K3 = Convert.ToDouble(.Text);//制热机组装机容量
                                      //txtOutput15
            Double K4 = T5 / 10000 * T1 / 4 * T2 / 100.0;//制热量
            //Double P12 = Convert.ToDouble(txtInput16.Text);//单位热价（含税）
            //txtOutput16.Text 
            Double K5 = (T7 / (Math.Pow(10, 9) / 4182 / 860));//单位热价
                                                              //txtOutput14.Text
            Double K6 = K4 * K5;//2-热费
            //Double P13 = Convert.ToDouble(txtInput9.Text);//制冷机组装机容量
            //txtOutput18.Text 
            Double K7 = T9 / 10000 * T1 / 4 * 3 * T2 / 100;//制冷量
                                                           //Double P14 = Convert.ToDouble(txtInput18.Text);//单位冷价（含税）
                                                           //txtOutput19.Text
            Double K8 = (T8 / (Math.Pow(10, 9) / 4182 / 860));//单位冷价
            //txtOutput17.Text
            Double K9 = K7 * K8;//3-冷费
                                //txtOutput20.Text
            Double K10 = K9 + K2 + K6;
            return K10;//总收入

        }
        #endregion
        #region   总成本1+2+3+4+5
        //txtInput14  txtInput11 txtInput10  txtInput7 txtInput19  txtInput20  txtInput12  txtOutput27  txtOutput28 txtOutput3 txtInput22  txtOutput31 txtInput21
        public double AllCost(double N1, double N2, double N3, double N4, double N5, double N6, double N7, double N8, double N9, double N10, double N11, double N12, double N13)
        {

            //Double P1 = Convert.ToDouble(txtInput7.Text); //发电机组装机容量
            //Double P2 = Convert.ToDouble(txtInput12.Text);//单位电力装机投资
            //txtOutput1.Text = (P1 * P2 / 10000).ToString("0");//1-总投资
            Double D5 = N4 * N7 / 10000;
            //Double P3 = Convert.ToDouble(txtOutput3.Text);//单位补贴
            //txtOutput2.Text = (P3 * P1 / 10000).ToString("0");//补贴
            Double D6 = N10 * N4 / 10000;
            //Double P8 = Convert.ToDouble(txtInput10.Text);//年利用小时数
            //Double P9 = Convert.ToDouble(txtInput11.Text) / 100;//负荷率
            //txtOutput12.Text = (Convert.ToDouble(txtInput7.Text) * P8 / 10000 * P9).ToString(); //发电量
            Double D1 = N4 * N3 / 10000 * N2 / 100;
            //Double P15 = Convert.ToDouble(txtInput14.Text);//发电气耗
            //txtOutput23.Text = (P15 * P9 * Convert.ToDouble(txtOutput12.Text)).ToString();//用气量
            Double D2 = N1 * N2 / 100 * D1;
            //txtOutput21.Text = (Convert.ToDouble(txtOutput23.Text) * Convert.ToDouble(txtOutput22.Text)).ToString();//1-燃料成本
            Double D3 = D2 * N5;
            //txtOutput25.Text = txtInput20.Text;//折旧年限
            //txtOutput24.Text = (Convert.ToDouble(txtOutput1.Text) / Convert.ToDouble(txtOutput25.Text)).ToString();//2-折旧
            Double D4 = D5 / N6;
            //txtOutput4.Text = (Convert.ToDouble(txtOutput1.Text) - Convert.ToDouble(txtOutput2.Text)).ToString("0"); //总投资（去除补贴）
            Double D7 = D5 - D6;
            //Double P16 = Convert.ToDouble(txtOutput27.Text);//贷款比例
            //Double P17 = Convert.ToDouble(txtOutput28.Text);//贷款利率
            //txtOutput26.Text = (Convert.ToDouble(txtOutput4.Text) * P16 * P17).ToString();//3-财务成本
            Double D10 = D7 * N9 * N8;
            //txtOutput30.Text = txtInput22.Text;//人员数
            //txtOutput29.Text = (Convert.ToDouble(txtInput22.Text) * Convert.ToDouble(txtOutput31.Text)).ToString();//4-人工成本
            Double D8 = N11 * N12;
            //txtOutput33.Text = txtInput21.Text;//单位运维成本
            //txtOutput32.Text = (Convert.ToDouble(txtOutput12.Text) * Convert.ToDouble(txtOutput33.Text)).ToString();//5-运维等成本
            Double D9 = D1 * N13;
            //总成本
            //txtOutput34.Text = (Convert.ToDouble(txtOutput21.Text) + Convert.ToDouble(txtOutput24.Text) + Convert.ToDouble(txtOutput26.Text) + Convert.ToDouble(txtOutput29.Text) + Convert.ToDouble(txtOutput32.Text)).ToString();
            Double D11 = D3 + D4 + D10 + D8 + D9;
            return D11;
        }
        #endregion
}
}
