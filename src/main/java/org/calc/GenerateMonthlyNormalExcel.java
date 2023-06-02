package org.calc;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.RoundingMode;
import java.text.DecimalFormat;

public class GenerateMonthlyNormalExcel {
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
        // 获取到最终的文件
        File destnation_file = new File("C:\\Users\\84334\\Desktop\\order\\2023\\票\\月度销售统计\\正常价格\\4月幼乐鲜.xlsx");
        FileInputStream destnation_file_stream = new FileInputStream(destnation_file);
        XSSFWorkbook workbook = new XSSFWorkbook(destnation_file_stream);
        // sheet操作结果页
        XSSFSheet result_sheet = workbook.getSheetAt(0);



        //
        handle_daily_nromal(4, result_sheet);
        FileOutputStream fileOutputStream = new FileOutputStream(destnation_file);
        workbook.write(fileOutputStream);
    }


    private static void handle_daily_nromal(int require_month, XSSFSheet result_sheet) throws Exception {
        File source_file = new File("C:\\Users\\84334\\Desktop\\order\\2023\\票\\单子综合");
        File[] files = source_file.listFiles();
        for (File file : files) {
//            System.out.println(file.getName());
            String name = file.getName();
            String[] split = name.split("\\.");
            int month = Integer.valueOf(split[0]);
            int day = Integer.valueOf(split[1]);
            // 如果到了超出的月份，运行结束
            if (month > require_month) break;
            // 假如不是指定月份，不看
            if (month != require_month) continue;
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

    /**
     * 根据当前操作的文件和页号来确定应该操作最终那个大文件的哪一列
     * @param real_file 此时在操作的某个文件
     * @param sheet_number  此时操作的某个文件的某页
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
