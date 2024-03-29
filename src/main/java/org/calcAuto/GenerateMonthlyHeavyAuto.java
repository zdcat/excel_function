package org.calcAuto;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.Iterator;

public class GenerateMonthlyHeavyAuto {
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

    private static BigDecimal watermelonAll = new BigDecimal(0);
    private static BigDecimal watermelonWuYueZanCheng = new BigDecimal(0);

    private static boolean watermelon_flag = false;

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

    public static void main(String[] args) throws Exception {
        int m = Integer.parseInt(args[0]);

        clearAll(m);
        // 获取到最终的文件
        int day = Integer.parseInt(args[1]);
        setDailyHeavy(m, day);
        // 从第2行到第32行每行都计算各家当天各自总和
        setDailySumPerZooAndPerRow(m);
        // 从第2列到第21列计算每家各自每个月送了多少
        setDailySumPerZooAndPerColumn(m);
    }


    private static void setDailySumPerZooAndPerColumn(int m) throws Exception {
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2024\\票\\月度销售统计\\重量\\" + m + "月幼乐鲜重量.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);

        // 只操作第32行
        Row row = result_sheet.getRow(32);
        BigDecimal sumPerZoo = new BigDecimal(0);
        BigDecimal sumFruit = new BigDecimal(0);
        BigDecimal sumVege = new BigDecimal(0);
        for (Cell cell : row) {
            int columnIndex = cell.getColumnIndex();
            if (columnIndex == 0) {
                continue;
            }
            if (columnIndex == 5 || columnIndex == 10 || columnIndex == 15 || columnIndex == 20) {
                cell.setCellValue(sumPerZoo.toString());
                sumPerZoo = new BigDecimal(0);
                continue;
            }
            if (columnIndex == 21) {
                break;
            }

            BigDecimal sum = new BigDecimal(0);
            for (int i = 1; i <= 31; i++) {
                BigDecimal num = new BigDecimal(get_cell_value(result_sheet.getRow(i), columnIndex))
                        .setScale(2, BigDecimal.ROUND_HALF_UP);
                sum = sum.add(num);
            }

            cell.setCellValue(sum.toString());
            sumPerZoo = sumPerZoo.add(sum);
            // 偶数
            if (columnIndex % 5 % 2 == 0) {
                sumFruit = sumFruit.add(sum);
//                System.out.println(sum);
            } else {
                sumVege = sumVege.add(sum);
//                System.out.println("    " + sum);
            }
        }

        // 拿到蔬菜和水果各自的斤数
//        sumFruit = format_double(sumFruit);
//        sumVege = format_double(sumVege);
        // 写入蔬菜水果各自斤数
        XSSFCell sumFruitCell = result_sheet.getRow(33).getCell(2);
        XSSFCell sumVegeCell = result_sheet.getRow(34).getCell(2);
        sumFruitCell.setCellValue(sumFruit.toString());
        sumVegeCell.setCellValue(sumVege.toString());

        // 写入西瓜总数量
        XSSFCell sumAllWatermelon = result_sheet.getRow(35).getCell(2);
        sumAllWatermelon.setCellValue(watermelonAll.toString());
        // 写入吾悦赞城西瓜总数量
        XSSFCell sumWuYueZanChengWatermelon = result_sheet.getRow(36).getCell(2);
        sumWuYueZanChengWatermelon.setCellValue(watermelonWuYueZanCheng.toString());


        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }


    private static void setDailySumPerZooAndPerRow(int m) throws Exception {
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2024\\票\\月度销售统计\\重量\\" + m + "月幼乐鲜重量.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);

        // 记录当月总销售

        for (Row row : result_sheet) {
            // 第一行不考虑
            if (row.getRowNum() == 0) {
                continue;
            }
            // 只计算第2行到第32行
            if (row.getRowNum() == 32) {
                break;
            }


            for (int times = 0; times < 4; times++) {
                BigDecimal sum = new BigDecimal(0)
                        .setScale(2, BigDecimal.ROUND_HALF_UP);
                for (int i = 1 + times * 5; i <= 4 + times * 5; i++) {
                    BigDecimal num = new BigDecimal(get_cell_value(row, i))
                            .setScale(2, BigDecimal.ROUND_HALF_UP);
                    // 更新某一家当天总销售
                    sum = sum.add(num);
                }

                // 设置某一家当天总销售
                Cell cell = row.getCell(5 + times * 5);
                cell.setCellValue(sum.toString());
            }
            // 给当天总销售设置
            Cell sumDailyCell = row.getCell(21);

        }


        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }


    private static void clearAll(int m) throws Exception {
        // 获取到最终的文件
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2024\\票\\月度销售统计\\重量\\" + m + "月幼乐鲜重量.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);


        for (Row row : result_sheet) {
            // 第一行不处理
            if (row.getRowNum() == 0) {
                continue;
            }
            if (row.getRowNum() == 34) {
                break;
            }

            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getColumnIndex() == 0) {
                    continue;
                }
                cell.setCellValue(0);
            }

        }

        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }

    private static void setDailyHeavy(int m, int day) throws Exception {
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2024\\票\\月度销售统计\\重量\\" + m + "月幼乐鲜重量.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);


        handle_daily_nromal(m, day, result_sheet);
        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);

        System.out.println("西瓜重量：" + watermelonAll);
        System.out.println("五月赞城" + watermelonWuYueZanCheng);
    }


    private static void handle_daily_nromal(int require_month, int require_day, XSSFSheet result_sheet) throws Exception {
        String userName = System.getProperty("user.name");
        File source_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2024\\票\\单子综合");
        File[] files = source_file.listFiles();
        for (File file : files) {
//            System.out.println(file.getName());
            String name = file.getName();
            String[] split = name.split("\\.");
            int month = Integer.valueOf(split[0]);
            int day = Integer.valueOf(split[1]);

            // 假如不是指定月份，不看
            if (month != require_month) continue;
            if (day > require_day) continue;
            System.out.println(month + "月" + day);

            // 指定月的每个文件夹的绝对路径
            String abso_path = file.getAbsolutePath();
            // 指定月的每天的正常价格的目录的绝对路径
            // C:\Users\84334\Desktop\order\2024\票\单子综合\4.10\正常价格
            String real_abso_path = abso_path + "\\正常价格";
//            System.out.println(real_abso_path);

            File real_file_per_day = new File(real_abso_path);
            File[] real_files = real_file_per_day.listFiles();
            // 现在的realFile就是真正的某个month的某day的正常价格的单子，我们要解析的其实就是这张单子
            for (File real_file : real_files) {
                XSSFWorkbook sheets = new XSSFWorkbook(new FileInputStream(real_file));


                if (real_file.getName().contains("吾") || real_file.getName().contains("赞")) {
                    watermelon_flag = true;
                }
//                if (day < 10)  watermelon_flag = false;

                for (int sheet_number = 0; sheet_number < 4; sheet_number++) {
                    // 拿到这一页
                    XSSFSheet sheet = sheets.getSheetAt(sheet_number);
//                    System.out.println(real_file.getName());
//                    System.out.println("第" + sheet_number + "页");
                    // 获得这一页的总和,并且格式化
                    BigDecimal per_sheet_value = get_per_sheet_value(sheet);

//                    System.out.println(per_sheet_value);

                    // row为到时候修改的第几行，比如4月2号，修改的是第3行,col为修改的第几列
                    int row_num = day;
                    int col_num = get_col_num(real_file, sheet_number);

                    XSSFRow result_row = result_sheet.getRow(row_num);

                    String cellValue = get_cell_value(result_row, col_num);

                    // 写入这一页的总和
                    if (cellValue.equals("")) {
                        result_row.getCell(col_num).setCellValue(per_sheet_value.toString());
                    } else {
                        BigDecimal old_row_value = new BigDecimal(get_cell_value(result_row, col_num))
                                .setScale(2, BigDecimal.ROUND_HALF_UP);
                        BigDecimal new_row_value = old_row_value.add(per_sheet_value)
                                .setScale(2, BigDecimal.ROUND_HALF_UP);
                        result_row.getCell(col_num).setCellValue(new_row_value.toString());
                    }

                }


                watermelon_flag = false;
            }


        }

    }

    /**
     * 根据当前操作的文件和页号来确定应该操作最终那个大文件的哪一列
     *
     * @param real_file    此时在操作的某个文件
     * @param sheet_number 此时操作的某个文件的某页
     * @return
     */
    private static int get_col_num(File real_file, int sheet_number) {
        String file_name = real_file.getName().substring(0, 1);
        int col = 0;
        if (file_name.equals("吾")) {
            col = 1 + sheet_number;
        } else if (file_name.equals("兴")) {
            col = 6 + sheet_number;
        } else if (file_name.equals("圪")) {
            col = 11 + sheet_number;
        } else if (file_name.equals("赞")) {
            col = 16 + sheet_number;
        }
        return col;
    }

    /**
     * 获得每页的总和，sum为初始值
     *
     * @param sheet 当前操作的某页
     *
     */
    private static BigDecimal get_per_sheet_value(Sheet sheet) {
        int i = 0;
        BigDecimal sum = new BigDecimal(0);
        for (Row row : sheet) {
            if (i <= 4) {
                i++;
                continue;
            }
            int status = verify(row);
            // null行直接break
            if (status == ROW_NULL) break;
            // 介于数字和合计行的不计算
            if (status == ROW_NUM_END) continue;
            // 如果到了合计行
            if (status == ROW_SUM) {
                // 更新合计金额
                break;
            }
            DataFormatter formatter = new DataFormatter();
            Cell cell2 = row.getCell(1);
            String s2 = formatter.formatCellValue(cell2);

            BigDecimal quantity = new BigDecimal(get_cell_value(row, 4))
                    .setScale(2, BigDecimal.ROUND_HALF_UP);
            if (s2.contains("西瓜") && watermelon_flag) {
                watermelonWuYueZanCheng = watermelonWuYueZanCheng.add(quantity)
                        .setScale(2, BigDecimal.ROUND_HALF_UP);
            }
            if (s2.contains("西瓜")) {
                watermelonAll = watermelonAll.add(quantity)
                        .setScale(2, BigDecimal.ROUND_HALF_UP);
            }

            sum = sum.add(quantity);
        }
        return sum;
    }

    private static Double format_double(Double result) {
        DecimalFormat df = new DecimalFormat("#.#");
        df.setRoundingMode(RoundingMode.HALF_UP); // 设置四舍五入规则
        String formattedNumber = df.format(result);
        return Double.parseDouble(formattedNumber);
    }

    private static String get_cell_value(Row row, int index) {
        DataFormatter dataFormatter = new DataFormatter();
        String s = dataFormatter.formatCellValue(row.getCell(index));
        return s;
    }
}



