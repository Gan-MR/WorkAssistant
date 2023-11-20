using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace 工作助手.Function
{
    class characterManipulation
    {
        //字符转换功能开始
        //GOGOGO-----------------------------------------------
        public string numtoUpper(int num)
        {
            string str = num.ToString();
            string rstr = "";
            int n;
            for (int i = 0; i < str.Length; i++)
            {
                n = Convert.ToInt16(str[i].ToString()); // char转数字,转换为字符串，再转数字
                switch (n)
                {
                    case 0:
                        rstr = rstr + "零";
                        break;
                    case 1:
                        rstr = rstr + "一";
                        break;
                    case 2:
                        rstr = rstr + "二";
                        break;
                    case 3:
                        rstr = rstr + "三";
                        break;
                    case 4:
                        rstr = rstr + "四";
                        break;
                    case 5:
                        rstr = rstr + "五";
                        break;
                    case 6:
                        rstr = rstr + "六";
                        break;
                    case 7:
                        rstr = rstr + "七";
                        break;
                    case 8:
                        rstr = rstr + "八";
                        break;
                    default:
                        rstr = rstr + "九";
                        break;
                }
            }
            return rstr;
        }

        // 月转化为大写
        public string monthtoUpper(int month)
        {
            if (month < 10)
            {
                return numtoUpper(month);
            }
            else if (month == 10)
            {
                return "十";
            }
            else
            {
                return "十" + numtoUpper(month - 10);
            }
        }

        // 日转化为大写
        public string daytoUpper(int day)
        {
            if (day < 20)
            {
                return monthtoUpper(day);
            }
            else
            {
                string str = day.ToString();
                if (str[1] == '0')
                {
                    return numtoUpper(Convert.ToInt16(str[0].ToString())) + "十";
                }
                else
                {
                    return numtoUpper(Convert.ToInt16(str[0].ToString())) + "十"
                        + numtoUpper(Convert.ToInt16(str[1].ToString()));
                }
            }
        }

        // 日期转换为大写
        public string dateToUpper(DateTime date)
        {
            int year = date.Year;
            int month = date.Month;
            int day = date.Day;
            return numtoUpper(year) + "年" + monthtoUpper(month) + "月" + daytoUpper(day) + "日";
        }

        // 从文本中找出数字日期并转换为大写
        public string ConvertDatesInText(string text)
        {
            string pattern = @"\d{4}年\d{1,2}月\d{1,2}日|\d{4}年\d{1,2}月|\d{4}年|\d{3}年\d{1,2}月\d{1,2}日|\d{3}年\d{1,2}月|\d{3}年";
            MatchEvaluator evaluator = new MatchEvaluator(ConvertDateMatchEvaluator);
            string output = Regex.Replace(text, pattern, evaluator);
            return output;
        }
        //关键代码^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


        private string ConvertDateMatchEvaluator(Match match)
        {
            DateTime date;
            if (DateTime.TryParse(match.Value, out date))
            {
                if (match.Value.EndsWith("年"))
                {
                    if (date.Year < 1000)
                    {
                        return numtoUpper(date.Year) + "年";
                    }
                    else
                    {
                        return numtoUpper(date.Year) + "年";
                    }
                }
                else if (match.Value.EndsWith("月"))
                {
                    if (date.Year < 1000)
                    {
                        return numtoUpper(date.Year) + "年" + monthtoUpper(date.Month) + "月";
                    }
                    else
                    {
                        return numtoUpper(date.Year) + "年" + monthtoUpper(date.Month) + "月";
                    }
                }
                else if (match.Value.EndsWith("日"))
                {
                    return dateToUpper(date);
                }
                else
                {
                    return dateToUpper(date);
                }
            }
            else
            {
                return match.Value;
            }
        }

        //字符转换功能结束
        //ENDENDEND-----------------------------------------------

        //括号及字符清理
        public static string RemoveParentheses(string input)
        {
            string pattern = @"\(.*?\)|（.*?）";
            string output = Regex.Replace(input, pattern, "");
            return output;
        }
        public static string RemoveSpecialSymbols(string input,string substitution)
        {
            string rt= input.Replace("~", "至");// 将~替换成“至”
            rt= Regex.Replace(rt, ":", "："); // 冒号
            rt = Regex.Replace(rt, ",", "，");// 逗号
            rt = Regex.Replace(rt, @"\?", "？");// 问号
            rt = Regex.Replace(rt, "!", "！");// 感叹号
            rt = Regex.Replace(rt, "；", "。");//中文分号转句号
            rt = Regex.Replace(rt, ";", "。");//英文分号转句号
            rt = Regex.Replace(rt, " ", "");//去除空格
            rt = Regex.Replace(rt, @"(?<!\d)\.(?!\d)", "。");//前后无数字的.替换。 
            rt = Regex.Replace(rt, @"["+substitution+"]+", "");// 删除特殊符号
            return rt;
        }
    }
}
