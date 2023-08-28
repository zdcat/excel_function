package org.test;

import java.math.BigDecimal;
import java.text.DecimalFormat;

public class DoubleToUpperCase {
    public static void main(String[] args) {
        double number = 12345.67;

        // 使用BigDecimal将double转换为字符串，并设置保留两位小数
        String numberString = new BigDecimal(number).setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();

        // 使用DecimalFormat将字符串格式化为大写形式
        DecimalFormat decimalFormat = new DecimalFormat("###,###.00");
        decimalFormat.setParseBigDecimal(true);
        try {
            BigDecimal parsedNumber = (BigDecimal) decimalFormat.parse(numberString);
            String upperCaseNumber = parsedNumber.toPlainString().toUpperCase();
            System.out.println(upperCaseNumber);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
