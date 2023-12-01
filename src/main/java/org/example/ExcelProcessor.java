package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.test.DecimalToChinese;
import org.test.DoubleToChinese;

public class ExcelProcessor {
    /**
     * 表明该row为要进行计算的数字行
     */
    private static final int ROW_NUM = 1;
    /**
     * 表明该行是合计那一行，在这一行要更新sum值
     */
    private static final int ROW_SUM = 2;

    /**
     * 表明到了最底行，最低行为null
     */
    private static final int ROW_NULL = 3;

    /**
     * 这种行介于最低的要计算的数字行和合计行之间的空白行
     */
    private static final int ROW_NUM_END = 4;


    public static void main(String[] args) throws Exception {
        // 赋值产生加3元文件，有的话删除
        String date = args[0];
        // 传入日期，进行加价文件的复制
        CopyFolderExample.main(args);
        String userName = System.getProperty("user.name");
        File folder = new File("C:\\Users\\" + userName +
                "\\Desktop\\order\\2023\\票\\单子综合\\" + args[0] + "\\" + getFileName(args) + " 加3元");
        File[] listOfFiles = folder.listFiles();

        // map的key为页号，value为这一页要加的加价，一般蔬菜1.5，水果3
        HashMap<Integer, Double> map = new HashMap<>();
        map.put(0, Double.valueOf(args[2]));
        map.put(1, Double.valueOf(args[1]));
        map.put(2, Double.valueOf(args[2]));
        map.put(3, Double.valueOf(args[1]));

        for (File file : listOfFiles) {
            if (file.isFile()) {
                String name = file.getName();
                for (int i = 3; i < args.length; i++) {
                    // 勾选了加价的某家幼儿园才会进行加价
                    if (name.contains(args[i])) {
                        handle_file(file.getAbsolutePath(), map);
                    }
                }
            }
        }

    }

    private static String getFileName(String[] args) {
        int len = args.length;
        String name = "";
        for (int i = 3; i < len; i++) {
            if (i != len - 1) {
                name += args[i];
                name += "-";
            } else {
                name += args[i];
            }
        }
        return name;
    }

    private static void handle_file(String fileName, HashMap<Integer, Double> map) {
        try (FileInputStream file = new FileInputStream(fileName);
             XSSFWorkbook workbook = new XSSFWorkbook(file)) {

            // i为页号
            for (int i = 0; i < 4; i++) {
                Sheet sheet = workbook.getSheetAt(i); // 获取第一个工作表
                int times = 0;
                // sum为这一页的新值
                BigDecimal decimal_sum = BigDecimal.valueOf(0);

                for (Row row : sheet) {
                    if (times <= 4) {
                        times++;
                        continue;
                    }

                    // 先判断这一行属于什么类型
                    int status = verify(row);
                    // null行直接break
                    if (status == ROW_NULL) break;
                    // 介于数字和合计行的不计算
                    if (status == ROW_NUM_END) continue;
                    // 如果到了合计行
                    if (status == ROW_SUM) {
                        // 更新合计金额
                        Cell cell = row.getCell(6);
                        // 写入这一页的总的新值
                        cell.setCellValue(decimal_sum.toString());

                        // 写入这一页总的新值的中文
                        Cell cell_chinese_value = row.getCell(2);
                        String chineseString = DecimalToChinese.getChineseString(decimal_sum);
                        cell_chinese_value.setCellValue(chineseString);
                        break;
                    }

                    // 不断遍历每一行，不断的加上每一行的新值
                    decimal_sum = decimal_sum.add(calc(row, map.get(i)))
                            .setScale(2, BigDecimal.ROUND_HALF_UP);

                }

            }
            System.out.println(fileName + "修改完毕");

            try (FileOutputStream outFile = new FileOutputStream(fileName)) {
                workbook.write(outFile);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 生成该row的新值，amount为这一页要加的新值
    private static BigDecimal calc(Row row, Double amount) {
        DataFormatter dataFormatter = new DataFormatter();
        // 设置新价格
        Cell cell_price = row.getCell(5);
        Cell cell_water = row.getCell(1);
        String cell_name = dataFormatter.formatCellValue(cell_water);
        String s_price = dataFormatter.formatCellValue(cell_price);

        // 获得旧价格
        BigDecimal decimal_old_price = new BigDecimal(s_price)
                .setScale(2, BigDecimal.ROUND_HALF_UP);
        // 获得新价格
        BigDecimal decimal_new_price = new BigDecimal(0).add(decimal_old_price)
                .setScale(2, BigDecimal.ROUND_HALF_UP);

        // 西瓜目前加1元
        if (cell_name.contains("西瓜")) {
            System.out.println("西瓜");
            decimal_new_price.subtract(new BigDecimal(2))
                    .setScale(2, BigDecimal.ROUND_HALF_UP);
        }
        // 写入新价格
        cell_price.setCellValue(decimal_new_price.toString());


        // 获得数量
        Cell cell_quantity = row.getCell(4);
        String s_quantity = dataFormatter.formatCellValue(cell_quantity);
        BigDecimal decimal_quantity = new BigDecimal(s_quantity)
                .setScale(2, BigDecimal.ROUND_HALF_UP);

        // 设置新金额
        // 得到此行新金额
        BigDecimal decimal_result = decimal_new_price.multiply(decimal_quantity)
                .setScale(2, BigDecimal.ROUND_HALF_UP);
        // 写入新金额
        Cell cell_result = row.getCell(6);
        cell_result.setCellValue(decimal_result.toString());
        return decimal_result;
    }

    private static int verify(Row row) {

        // 假如某一行的第一列是null，说明所有的行走完了
        Cell cell1 = row.getCell(0);
        if (cell1 == null) {
            return ROW_NULL;
        }

        // 假如到了合计行，说明加完了
        DataFormatter formatter = new DataFormatter();
        String s1 = formatter.formatCellValue(cell1);
        if (s1.equals("合计")) {
            return ROW_SUM;
        }

        // 如果走到了数字行的末尾，那也要说明加完了
        Cell cell2 = row.getCell(1);
        String s2 = formatter.formatCellValue(cell2);
        if (s2.equals("")) {
            return ROW_NUM_END;
        }

        // 走到这里说明这一行是真正的数字行
        return ROW_NUM;
    }
}
