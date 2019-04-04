using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 天然气市场需求分析软件_求你不死机版_
{
    class Common
    {
        public static string path = null;
        #region   这是一个我自己定义的区域，目的是为了是代码结构变得简单，分类作用  锅炉规模检测和年用气小时数
        //锅炉规模检测
        public static void ParameterErrorDetectionBoilerScale(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //年用气小时数检测
        public static void ParameterErrorDetectionGasTime(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //城市类型检测
        public static void ParameterErrorDetectionCitySize(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新选择。");
                
            }
        }
        //城市人口数量检测
        public static void ParameterErrorDetectionCityPopulation(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");
                
            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,100000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //已气化居民数量检测
        public static void ParameterErrorDetectionAlreadyGasUsedPopulation(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,100000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //居民用气量检测
        public static void ParameterErrorDetectionNowResidentGasUsed(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,100000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //现状采暖面积
        public static void ParameterErrorDetectionNowHeatingArea(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,100000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //现状年份
        public static void ParameterErrorDetectionNowYear(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]
            if (targetValue1 <2017)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}不在参数范围（2017，2035]之内，请重新输入。");
            }

            if (targetValue1 > 2035)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}不在参数范围（2017，2035]之内，请重新输入。");
            }

        }
        //产品产量检测
        public static void ParameterErrorDetectionProductAmount(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]
           

            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //用气定额检测
        public static void ParameterErrorDetectionGasUsedAmount(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]
            

            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //机组数量检测
        public  static void ParameterErrorDetectionMachineAmount(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");


            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //单台机组容量检测
        public static void ParameterErrorDetectionSingleMachineValume(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            //if (targetValue1 > 10000000)
            //{
            //    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            //}

        }
        //机组效率检测
        public static void ParameterErrorDetectionMachineEfficiency(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]
            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //年利用小时数检测
        public static void ParameterErrorDetectionAnnualUsedhours(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
        }
        //天然气热值检测
        public static void ParameterErrorDetectionGasHeating(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //年发电量检测
        public static void ParameterErrorDetectionAnnualElectrityProduct(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //年供热量
        public static void ParameterErrorDetectionAnnualHeatingProduct(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //年制冷量检测
        public static void ParameterErrorDetectionAnnualColdProduct(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");
            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //天然气热值量检测
        public static void ParameterErrorDetectionNaturalGasHeatValue(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //热电联产三联供机组数量检测
        public static void ParameterErrorDetectionMachineAmount2(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （2018,2035]


            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        public static void ParameterErrorDetectionSingleMachineValume2(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0，10000000]
            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        public static void ParameterErrorDetectionMachineEfficiency2(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            ////参数检测，判断输入是否在规定范围内 （2018,2035]
            //if (targetValue1 > 10000000)
            //{
            //    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            //}

        }
        //化工产品产量检测
        public static void ParameterErrorDetectionProductOutput(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            ////参数检测，判断输入是否在规定范围内 （0，10000000]
            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //化工用气定额检测
        public static void ParameterErrorDetectionGasUsed(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            ////参数检测，判断输入是否在规定范围内 （0，10000000]
            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        public static void ParameterErrorDetectionNaturalGasConsume(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            ////参数检测，判断输入是否在规定范围内 （0，10000000]
            if (targetValue1 > 10000000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000000]，请重新输入。");
            }

        }
        //发电机组装机容量检测
        public static void ParameterErrorDetectionDynamoValume(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
           
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //制热机组装机容量检测
        public static void ParameterErrorDetectionHeatProductValume(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
           
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //制冷机组装机容量检测
        public static void ParameterErrorDetectionColdProductValume(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //分布式能源年利用小时数检测
        public static void ParameterErrorDetectionAnnualUsedHour(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //负荷率检测
        public static void ParameterErrorDetectionBurthenRatio(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //单位电力组装投资检测
        public static void ParameterErrorDetectionSingleElectrityInvest(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //单位装机补贴检测
        public static void ParameterErrorDetectionSingleAllowance(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //发电气耗检测
        public static void ParameterErrorDetectionElectrityProductGasUsed(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //供热气耗检测
        public static void ParameterErrorDetectionHeatingProductGasUsed(string str1, string str2)
        {
            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");
            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");
            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //单位电价检测
        public static void ParameterErrorDetectionSingleElectrityPrice(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //单位热价检测
        public static void ParameterErrorDetectionPerHeaatingPrice(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //单位冷价检测
        public static void ParameterErrorDetectionPerColdPrice(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //单位气价检测
        public static void ParameterErrorDetectionPerGasPrice(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }
        //折旧年限检测
        public static void ParameterErrorDetection(string str1, string str2)
        {

            //参数检测，判断txtInput2输入是否为空
            if (str2 == "")
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为空，请重新输入。");

            }
            //参数检测，判断txtInput2输入是否含有字符
            foreach (char c in str2)
            {
                if (char.IsLetter(c))
                {
                    throw new InvalidOperationException("输入参数{" + str1 + str2 + "}输入参数含有字符，请重新输入。");
                }
            }
            double targetValue1 = Convert.ToDouble(str2);
            //参数检测，判断输入是否为数字、零
            if (targetValue1 < 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为负数，请重新输入。");

            }
            if (targetValue1 == 0)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}为零，请重新输入。");

            }
            //参数检测，判断输入是否在规定范围内 （0,1000000]
            if (targetValue1 > 100000)
            {
                throw new InvalidOperationException("输入参数{" + str1 + str2 + "}超过输入参数范围（0，10000]，请重新输入。");
            }

        }



        #endregion

    }
}
