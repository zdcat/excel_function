package org.test;

import java.math.BigDecimal;

public class DecimalMultiplicationExample {

    public static void main(String[] args) {
        // 创建两个BigDecimal对象
        BigDecimal operand1 = new BigDecimal("12.345");
        BigDecimal operand2 = new BigDecimal("6.789");

        // 进行乘法计算
        BigDecimal result = operand1.multiply(operand2);

        // 设置小数保留位数为两位，使用ROUND_HALF_UP四舍五入
        result = result.setScale(2, BigDecimal.ROUND_HALF_UP);

        System.out.println("Result of multiplication: " + operand1);
        System.out.println("Result of multiplication: " + operand2);
        System.out.println("Result of multiplication: " + result);
    }
}
