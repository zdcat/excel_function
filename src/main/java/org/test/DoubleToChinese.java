package org.test;

public class DoubleToChinese {
    // 定义中文数字数组
    private static final String[] CHINESE_NUMBERS = {"零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"};

    // 定义中文单位数组
    private static final String[] CHINESE_UNITS = {"", "拾", "佰", "仟"};

    public static String getChineseString(double number){
        if (number == 0){
            return "零元整";
        }
        // 将double类型的数字转换为字符串，并去除小数点后多余的0
        String numberString = String.format("%.1f", number);
        if (numberString.endsWith(".0")) {
            numberString = numberString.substring(0, numberString.length() - 2);
        }

        // 如果存在小数点，则进行分割
        String integerPart;
        String decimalPart = "";
        if (numberString.contains(".")) {
            int decimalIndex = numberString.indexOf(".");
            integerPart = numberString.substring(0, decimalIndex);
            decimalPart = numberString.substring(decimalIndex + 1);
        } else {
            integerPart = numberString;
        }

        // 将整数部分转换为中文金额
        String chineseNumber = convertToChinese(integerPart) + "元";

        // 如果有小数部分，则将小数部分转换为中文金额并与整数部分拼接起来
        if (!decimalPart.isEmpty()) {
            chineseNumber += convertDecimalToChinese(decimalPart);
        }
        chineseNumber += "整";

        return chineseNumber;
    }
    public static void main(String[] args) {
        double number = 230.0;

        // 将double类型的数字转换为字符串，并去除小数点后多余的0
        String numberString = String.format("%.1f", number);
        if (numberString.endsWith(".0")) {
            numberString = numberString.substring(0, numberString.length() - 2);
        }

        // 如果存在小数点，则进行分割
        String integerPart;
        String decimalPart = "";
        if (numberString.contains(".")) {
            int decimalIndex = numberString.indexOf(".");
            integerPart = numberString.substring(0, decimalIndex);
            decimalPart = numberString.substring(decimalIndex + 1);
        } else {
            integerPart = numberString;
        }

        // 将整数部分转换为中文金额
        String chineseNumber = convertToChinese(integerPart) + "元";

        // 如果有小数部分，则将小数部分转换为中文金额并与整数部分拼接起来
        if (!decimalPart.isEmpty()) {
            chineseNumber += convertDecimalToChinese(decimalPart);
        }
        chineseNumber += "整";


        System.out.println(chineseNumber);

    }

    // 将数字字符串转换为中文
    private static String convertToChinese(String numberString) {
        StringBuilder result = new StringBuilder();
        int length = numberString.length();

        for (int i = 0; i < length; i++) {
            char c = numberString.charAt(i);
            int digit = Character.getNumericValue(c);

            // 如果当前位的数字不是零，则添加对应的中文数字和单位
            if (digit != 0) {
                result.append(CHINESE_NUMBERS[digit]);
                result.append(CHINESE_UNITS[length - i - 1]);
            } else {
                // 如果当前位的数字是零，且下一位数字不是零，则添加“零”
                if (i < length - 1 && numberString.charAt(i + 1) != '0') {
                    result.append("零");
                }
            }
        }

        return result.toString();
    }

    // 将小数部分转换为中文金额
    private static String convertDecimalToChinese(String decimalPart) {
        StringBuilder result = new StringBuilder("");

        char c = decimalPart.charAt(0);
        int digit = Character.getNumericValue(c);
        result.append(CHINESE_NUMBERS[digit]).append("角");

        return result.toString();
    }
}
