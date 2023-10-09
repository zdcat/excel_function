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
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.Iterator;

public class GenerateMonthlyNormalExcelAuto {
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
        // 先清除所有的数据
        clearAll(m);
        // 将每一天的数据填上去
        int day = Integer.parseInt(args[1]);
        setDailyNum(m, day);

        // 从第2行到第32行每行都计算各家当天各自总和，包括当天的总销售
        setDailySumPerZooAndPerRow(m);
        // 从第2列到第21列计算每家各自每个月送了多少
        setDailySumPerZooAndPerColumn(m);


    }

    private static void setDailySumPerZooAndPerColumn(int m) throws Exception {
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2023\\票\\月度销售统计\\正常价格\\" + m + "月幼乐鲜.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);

        // 只操作第32行
        Row row = result_sheet.getRow(32);
        double sumPerZoo = 0.0;
        for (Cell cell : row) {
            int columnIndex = cell.getColumnIndex();
            if (columnIndex == 0) {
                continue;
            }
            if (columnIndex == 5 || columnIndex == 10 || columnIndex == 15 || columnIndex == 20) {
                cell.setCellValue(sumPerZoo);
                sumPerZoo = 0.0;
                continue;
            }
            if (columnIndex == 21) {
                break;
            }

            double sum = 0;
            for (int i = 1; i <= 31; i++) {
                Double num = format_double(new Double(get_cell_value(result_sheet.getRow(i), columnIndex)));
                sum += num;
            }

            cell.setCellValue(sum);
            sumPerZoo += sum;
        }


        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }

    /**
     * 从第2行到第32行每行都计算各家当天各自总和，包括当天总销售以及当月总销售
     *
     * @param m
     * @throws Exception
     */
    private static void setDailySumPerZooAndPerRow(int m) throws Exception {
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2023\\票\\月度销售统计\\正常价格\\" + m + "月幼乐鲜.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);

        // 记录当月总销售
        double sumMonth = 0.0;
        for (Row row : result_sheet) {
            // 第一行不考虑
            if (row.getRowNum() == 0) {
                continue;
            }
            // 只计算第2行到第32行
            if (row.getRowNum() == 32) {
                break;
            }

            // 计算当天的总销售
            double sumDaily = 0.0;
            for (int times = 0; times < 4; times++) {
                double sum = 0.0;
                for (int i = 1 + times * 5; i <= 4 + times * 5; i++) {
                    Double num = format_double(new Double(get_cell_value(row, i)));
                    // 更新某一家当天总销售
                    sum += num;
                }

                // 设置某一家当天总销售
                Cell cell = row.getCell(5 + times * 5);
                cell.setCellValue(sum);
                // 更新当天总销售
                sumDaily += sum;
            }
            // 给当天总销售设置
            Cell sumDailyCell = row.getCell(21);
            sumDailyCell.setCellValue(sumDaily);
            // 更新当月总销售
            sumMonth += sumDaily;
        }

        // 给月度总和设置值
        Cell sumMonthCell = result_sheet.getRow(32).getCell(21);
        sumMonthCell.setCellValue(sumMonth);

        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }

    private static void setDailyNum(int m, int day) throws Exception {
        // 获取到最终的文件
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2023\\票\\月度销售统计\\正常价格\\" + m + "月幼乐鲜.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);


        //
        handle_daily_nromal(m, day, result_sheet);
        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }

    private static void clearAll(int m) throws Exception {
        // 获取到最终的文件
        String userName = System.getProperty("user.name");
        File destnation_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2023\\票\\月度销售统计\\正常价格\\" + m + "月幼乐鲜.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);


        for (Row row : result_sheet) {
            // 第一行不处理
            if (row.getRowNum() == 0) {
                continue;
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


    private static void handle_daily_nromal(int require_month, int require_day, XSSFSheet result_sheet) throws Exception {
        String userName = System.getProperty("user.name");
        File source_file = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2023\\票\\单子综合");
        File[] files = source_file.listFiles();
        for (File file : files) {
//            System.out.println(file.getName());
            String name = file.getName();
            String[] split = name.split("\\.");
            int month = Integer.valueOf(split[0]);
            int day = Integer.valueOf(split[1]);

            // 假如不是指定月份，不看
            if (month != require_month) continue;
            // 如果超过了某号就不看
            if (day > require_day) continue;
            System.out.println(month + "月" + day);

            // 指定月的每个文件夹的绝对路径
            String abso_path = file.getAbsolutePath();
            // 指定月的每天的正常价格的目录的绝对路径
            // C:\Users\84334\Desktop\order\2023\票\单子综合\4.10\正常价格
            String real_abso_path = abso_path + "\\正常价格";
//            System.out.println(real_abso_path);

            File real_file_per_day = new File(real_abso_path);
            File[] real_files = real_file_per_day.listFiles();
            // 现在的realFile就是真正的某个month的某day的正常价格的单子，我们要解析的其实就是这张单子
            for (File real_file : real_files) {
                XSSFWorkbook sheets = new XSSFWorkbook(new FileInputStream(real_file));

                for (int sheet_number = 0; sheet_number < 4; sheet_number++) {
                    // 拿到这一页
                    XSSFSheet sheet = sheets.getSheetAt(sheet_number);
//                    System.out.println(real_file.getName());
//                    System.out.println("第" + sheet_number + "页");
                    // 获得这一页的总和,并且格式化
                    Double per_sheet_value = format_double(get_per_sheet_value(sheet, 0.0));

//                    System.out.println(per_sheet_value);

                    // row为到时候修改的第几行，比如4月2号，修改的是第3行,col为修改的第几列
                    int row_num = day;
                    int col_num = get_col_num(real_file, sheet_number);

                    XSSFRow result_row = result_sheet.getRow(row_num);

                    String cellValue = get_cell_value(result_row, col_num);
                    if (cellValue.equals("")) {
                        result_row.getCell(col_num).setCellValue(per_sheet_value);
                    } else {
                        Double old_row_value = new Double(get_cell_value(result_row, col_num));
                        Double new_row_value = old_row_value + per_sheet_value;
                        result_row.getCell(col_num).setCellValue(new_row_value);
                    }

                }
            }


        }

    }

    private static String get_real_abso_path(String absoPath) throws Exception {
        File file = new File(absoPath);
        File[] files = file.listFiles();
        for (File file1 : files) {
            if (file1.isDirectory())
                if (file1.getName().contains("加")) {
                    return absoPath + "\\" + file1.getName();
                }
        }
        return "";
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
     * @param sum   初始值
     */
    private static Double get_per_sheet_value(Sheet sheet, Double sum) {
        int i = 0;
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

            Double quantity = new Double(get_cell_value(row, 4));
            Double price = new Double(get_cell_value(row, 5));
            Double per_row_value = new Double(quantity * price);
            per_row_value = format_double(per_row_value);
            sum += per_row_value;
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
