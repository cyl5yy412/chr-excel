package com.chryl.util;

import org.apache.commons.lang3.StringUtils;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;


public class CommonUtil {


    /**
     * 保留两位小数
     *
     * @param num
     * @return
     */
    public static String getTwoDecimal(double num) {
        DecimalFormat df = new DecimalFormat("######0.00");
        String result = df.format(num);
        return result;
    }

    public static double getTwoDecimalToDouble(double num) {
        double result = Double.parseDouble(getTwoDecimal(num));
        return result;
    }

    /**
     * 截取小数点之后的字符串  例：7.35-----7
     *
     * @param str
     * @return
     */
    public static String subString(String str) {
        String s = "";
        s = (str.contains(".") ? (str.substring(0, str.indexOf("."))) : str);
        return s;
    }

    /**
     * 处理为NaN和isInfinite情况的Rate
     *
     * @param rate
     * @return rate
     */
    public static double handleRate(double rate) {
        return (Double.isNaN(rate) ? (0) : (Double.isInfinite(rate) ? (1) : (rate)));
    }

    /**
     * @param exc_date
     * @return
     * @throws ParseException
     */
    public static String getUpMonth(String exc_date) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(sdf.parse(exc_date));
        calendar.add(Calendar.MONTH, -1);
        Date d = calendar.getTime();
        String mon = sdf.format(d);
        return mon;
    }

    public static String getExcDate() {
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        String date = df.format(new Date());
        return date;
    }

    ;

    public static String parseDate(String date) throws ParseException {
        SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        Date d = df.parse(date);
        return sdf.format(d);
    }

    public static String getMonth() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM");
        String mon = sdf.format(new Date());
        return mon;
    }

    public static String getFormatDate(Date date) {
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return df.format(date);
    }

    public static String getExcDate(String date) {
        String exc_date = date.substring(0, 4) + "-" + date.substring(4, 6) + '-' + date.substring(6, 8);
        return exc_date;
    }

    /**
     * 根据总分数得出档级
     *
     * @param allStore
     * @return
     */
    public static String computeGradelevel(int allStore) {
        if (allStore <= 10) {
            return "1";
        } else if (allStore <= 20) {
            return "2";
        } else if (allStore <= 30) {
            return "3";
        } else if (allStore <= 40) {
            return "4";
        } else if (allStore <= 50) {
            return "5";
        } else if (allStore <= 60) {
            return "6";
        } else if (allStore <= 70) {
            return "7";
        } else if (allStore <= 80) {
            return "8";
        } else if (allStore <= 90) {
            return "9";
        } else {
            return "10";
        }
    }

    public static int computeHonorstore(String honorlevel) {
        switch (honorlevel) {
            case "3"://国家级
                return 10;
            case "2"://省级
                return 7;
            case "1"://市级
                return 5;
            case "0"://县级
                return 2;
            default:
                return 0;
        }
    }

    /**
     * 计算学历分值
     *
     * @param education1
     * @param education2
     * @return
     */
    public static int computeEdustore(String education1, String education2) {
        if (StringUtils.isNotEmpty(education2) && StringUtils.isNotEmpty(education1)) {//如果第一学历第二学历都不为空
            if (Integer.parseInt(education2) > Integer.parseInt(education1)) {//第二学历高
                switch (education2) {
                    case "4"://博士研究生
                        return 23;
                    case "3"://硕士研究生
                        return 18;
                    case "2"://大学本科
                        return 13;
                    case "1"://专科
                        return 8;
                    default://中专（高中）及以下
                        return 5;
                }
            } else {//第一学历高或者两个学历一致
                switch (education1) {
                    case "4":
                        return 25;
                    case "3":
                        return 20;
                    case "2":
                        return 15;
                    case "1":
                        return 10;
                    default:
                        return 5;
                }
            }
        } else if (StringUtils.isNotEmpty(education2) && !StringUtils.isNotEmpty(education1)) {//第二学历不为空，第一学历为空
            switch (education2) {
                case "4"://博士研究生
                    return 23;
                case "3"://硕士研究生
                    return 18;
                case "2"://大学本科
                    return 13;
                case "1"://专科
                    return 8;
                default://中专（高中）及以下
                    return 5;
            }
        } else {//第一学历不为空，第二学历为空
            switch (education1) {
                case "4":
                    return 25;
                case "3":
                    return 20;
                case "2":
                    return 15;
                case "1":
                    return 10;
                default:
                    return 5;
            }
        }

    }

    /**
     * 计算工龄   薪酬年份-参加工作年份
     *
     * @return
     */
    public static int computeWorkYear(String workdate, String fixdate) {
        int yearNow = Integer.parseInt(fixdate.substring(0, 4));
        int yearWork = Integer.parseInt(workdate.substring(0, 4));
        int workyear = yearNow - yearWork;
        return workyear;
    }

    /**
     * 根据职务评定职级
     */
    public static String computeRankLevel(String post) {
        switch (post) {
            case "班子正职":
                return "13";
            case "班子副职":
                return "11";
            case "中层正职":
                return "8";
            case "中层副职":
                return "6";
            case "员工职级":
                return "4";
            case "见习期":
                return "1";
            default:
                return null;
        }

    }

    public static int computeProstore(String pro) {
        //String pro= (professionals.startsWith("高级"))?("3"):(professionals.startsWith("助理")?("1"):(professionals.startsWith("初级")?("0"):("2")));
        switch (pro) {
            case "3":
                return 25;
            case "2":
                return 20;
            case "1":
                return 15;
            case "0":
                return 10;
            default:
                break;
        }
        return 0;
    }

    /**
     * 校验固定薪酬表导入数据是否有效
     *
     * @param str
     * @return
     */
    public static boolean checkInput(String str) {
        boolean b = (null == str || "".equals(str)) ? (true) : (false);
        return b;
    }

    /**
     * 将yyyymm格式转为 yyyy-mm
     *
     * @param workdate
     * @return
     * @throws Exception
     */
    public static String toDate(String workdate) throws Exception {
        String date = "";
        if (StringUtils.isNotEmpty(workdate) && (workdate.length() == 6)) {
            DateFormat df = new SimpleDateFormat("yyyyMM");
            DateFormat df1 = new SimpleDateFormat("yyyy-MM");
            Date d;
            d = df.parse(workdate);
            date = df1.format(d);
        } else {
            throw new Exception("时间格式不正确，正确格式为yyyyMM，请修改后重新导入");
        }
        return date;
    }

    public static String toDate1(String workdate) throws Exception {
        String date = "";
        if (StringUtils.isNotEmpty(workdate) && (workdate.length() == 8)) {
            DateFormat df = new SimpleDateFormat("yyyyMMdd");
            DateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
            Date d;
            d = df.parse(workdate);
            date = df1.format(d);
        } else {
            throw new Exception("时间格式不正确，正确格式为yyyyMMdd，请修改后重新导入");
        }
        return date;
    }

    /**
     * 将百分数转为小数
     *
     * @param percent
     * @return
     */
    public static String percentToDecimal(String percent) {
        double decimal = Double.parseDouble(percent.replace("%", "")) / 100;
        return getTwoDecimal(decimal);

    }

    public static String getLastDayOfUpMonth(String fixdate) {
        DateFormat df = new SimpleDateFormat("yyyy-MM");
        DateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
        Calendar calendar = Calendar.getInstance();
        try {
            calendar.setTime(df.parse(fixdate));//设置日历实例时间
        } catch (ParseException e) {
            e.printStackTrace();
        }
        int year = calendar.get(Calendar.YEAR);
        int month = calendar.get(Calendar.MONTH);
        calendar.set(year, month, 0);
        return df1.format(calendar.getTime());
    }

    public static Map<String, String> getResultMap(List<Map<String, Object>> list) {
        Map<String, String> map = new HashMap<String, String>();
        String key;
        String value;
        for (Map<String, Object> item : list) {
            System.out.println(item);
            key = String.valueOf(item.get("KEY"));
            value = String.valueOf(item.get("VALUE"));
            map.put(key, value);
        }
        return map;
    }
}
