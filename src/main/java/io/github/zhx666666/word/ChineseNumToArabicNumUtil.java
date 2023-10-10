package io.github.zhx666666.word;

/**
 * 阿拉伯数字 和 中文汉字 相互转换工具类
 *
 * @author 明快de玄米61
 * @date 2022/9/13 15:40
 */
public class ChineseNumToArabicNumUtil {
    static char[] cnArr = new char [] {'零','一','二','三','四','五','六','七','八','九'};
    static char[] chArr = new char [] {'零','十','百','千','万','亿'};
    static String allChineseNum = "零一二三四五六七八九十百千万亿";
    static String allArabicNum = "0123456789";
    static String num1 = "一二三四五六七八九";
    static String num2 = "十百千万亿";
    static String zero = "零";

    /**
     * 将汉字中的数字转换为阿拉伯数字， 转换纯中文数字，
     * @param chineseNum
     * @return
     */
    public static int chineseNumToArabicNum(String chineseNum) {
        int result = 0;
        int temp = 1;//存放一个单位的数字如：十万
        int count = 0;//判断是否有chArr
        for (int i = 0; i < chineseNum.length(); i++) {
            boolean b = true;//判断是否是chArr
            char c = chineseNum.charAt(i);
            for (int j = 0; j < cnArr.length; j++) {
                if (c == cnArr[j]) {
                    if(0 != count){//添加下一个单位之前，先把上一个单位值添加到结果中
                        result += temp;
                        temp = 1;
                        count = 0;
                    }
                    // 下标，就是对应的值
                    temp = j;
                    b = false;
                    break;
                }
            }
            if(b){//单位{'十','百','千','万','亿'}
                for (int j = 0; j < chArr.length; j++) {
                    if (c == chArr[j]) {
                        switch (j) {
                            case 0:
                                temp *= 1;
                                break;
                            case 1:
                                temp *= 10;
                                break;
                            case 2:
                                temp *= 100;
                                break;
                            case 3:
                                temp *= 1000;
                                break;
                            case 4:
                                temp *= 10000;
                                break;
                            case 5:
                                temp *= 100000000;
                                break;
                            default:
                                break;
                        }
                        count++;
                    }
                }
            }
            if (i == chineseNum.length() - 1) {//遍历到最后一个字符
                result += temp;
            }
        }
        return result;
    }
    /**
     * 将字符串中的中文数字转换阿拉伯数字，其它非数字汉字不替换
     * @param chineseNum
     * @return
     */
    public static String chineseNumToArabicNumTwo(String chineseNum) {
        StringBuilder resultStr = new StringBuilder();
        int tempresult = 0;
        int temp = 1;//存放一个单位的数字如：十万
        int count = 0;//判断是否有单位
        // 重新将 temp, count, tempresult 设置为初始值
        boolean setInitial = false;
        // 以十百千万亿结束的在最后加
        boolean isAdd = false;
        boolean num1flag = false;
        boolean num2flag = false;
        for (int i = 0; i < chineseNum.length(); i++) {
            if (setInitial) {
                tempresult = 0;
                temp = 1;
                count = 0;
                setInitial = false;
            }
            boolean b = true;//判断是否是chArr
            char c = chineseNum.charAt(i);
            if (allChineseNum.indexOf(c) >= 0) {
                if (i < chineseNum.length() - 1 && num1.indexOf(c) >= 0 && num1.indexOf(chineseNum.charAt(i+1)) >= 0) {
                    num1flag = true;
                }
                for (int j = 0; j < cnArr.length; j++) {
                    if (c == cnArr[j]) {
                        if(0 != count){//添加下一个单位之前，先把上一个单位值添加到结果中
                            tempresult += temp;
                            temp = 1;
                            count = 0;
                        }
                        if (!isAdd && (i == chineseNum.length() - 1
                                || allChineseNum.indexOf(chineseNum.charAt(i+1)) < 0)) {
                            tempresult += j;
                            setInitial = true;
                            resultStr.append(tempresult);
                            isAdd = true;
                        }
                        // 下标+1，就是对应的值
                        temp = j;
                        b = false;
                        break;
                    }
                }
                if (num1flag) {
                    resultStr.append(temp);
                    num1flag = false;
                    setInitial = true;
                    continue;
                }

                boolean test = (i < chineseNum.length() - 1 && zero.indexOf(chineseNum.charAt(i+1)) >= 0 )
                        || (i >0 && zero.indexOf(chineseNum.charAt(i-1)) >= 0);
                if (i < chineseNum.length() - 1 && zero.indexOf(c) >= 0 && test ) {
                    num2flag = true;
                }
                if(b){//单位{'十','百','千','万','亿'}
                    for (int j = 0; j < chArr.length; j++) {
                        if (c == chArr[j]) {
                            switch (j) {
                                case 0:
                                    temp *= 1;
                                    break;
                                case 1:
                                    temp *= 10;
                                    break;
                                case 2:
                                    temp *= 100;
                                    break;
                                case 3:
                                    temp *= 1000;
                                    break;
                                case 4:
                                    temp *= 10000;
                                    break;
                                case 5:
                                    temp *= 100000000;
                                    break;
                                default:
                                    break;
                            }
                            count++;
                        }
                    }
                }
                if (num2flag) {
                    resultStr.append(temp);
                    num2flag = false;
                    setInitial = true;
                    continue;
                }
                if (!isAdd && (i == chineseNum.length() - 1
                        || allChineseNum.indexOf(chineseNum.charAt(i+1)) < 0)) {
                    tempresult += temp;
                    setInitial = true;
                    resultStr.append(tempresult);
                    isAdd = true;
                }
            } else {
                isAdd = false;
                resultStr.append(c);
            }
        }
        return resultStr.toString();
    }
    /**
     * 将数字转换为中文数字， 这里只写到了万
     * @param intInput
     * @return
     */
    public static String arabicNumToChineseNum(int intInput) {
        String si = String.valueOf(intInput);
        String sd = "";
        if (si.length() == 1) {
            if (intInput == 0) {
                return sd;
            }
            sd += cnArr[intInput];
            return sd;
        } else if (si.length() == 2) {
            if (si.substring(0, 1).equals("1")) {
                sd += "十";
                if (intInput % 10 == 0) {
                    return sd;
                }
            }
            else
                sd += (cnArr[intInput / 10] + "十");
            sd += arabicNumToChineseNum(intInput % 10);
        } else if (si.length() == 3) {
            sd += (cnArr[intInput / 100] + "百");
            if (String.valueOf(intInput % 100).length() < 2) {
                if (intInput % 100 == 0) {
                    return sd;
                }
                sd += "零";
            }
            sd += arabicNumToChineseNum(intInput % 100);
        } else if (si.length() == 4) {
            sd += (cnArr[intInput / 1000] + "千");
            if (String.valueOf(intInput % 1000).length() < 3) {
                if (intInput % 1000 == 0) {
                    return sd;
                }
                sd += "零";
            }
            sd += arabicNumToChineseNum(intInput % 1000);
        } else if (si.length() == 5) {
            sd += (cnArr[intInput / 10000] + "万");
            if (String.valueOf(intInput % 10000).length() < 4) {
                if (intInput % 10000 == 0) {
                    return sd;
                }
                sd += "零";
            }
            sd += arabicNumToChineseNum(intInput % 10000);
        }

        return sd;
    }

    /**
     * 判断传入的字符串是否全是汉字数字
     * @param chineseStr
     * @return
     */
    public static boolean isChineseNum(String chineseStr) {
        char [] ch = chineseStr.toCharArray();
        for (char c : ch) {
            if (!allChineseNum.contains(String.valueOf(c))) {
                return false;
            }
        }
        return true;
    }

    /**
     * 判断数字字符串是否是整数字符串
     * @param str
     * @return
     */
    public static boolean isNum(String str) {
        String reg = "[0-9]+";
        return str.matches(reg);
    }
}

