package org.useless;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Useless {
    /**
     * 把目录下的所有excel的送货单改成送货单（幼儿）或者送货单（老师）
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        String userName = System.getProperty("user.name");
        File files = new File("C:\\Users\\84334\\Desktop\\order\\2023\\票\\补单\\3月补单\\赞城3月");
        for (File file : files.listFiles()) {
            System.out.println(file.getName());
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
            for (int i = 0; i < 4; i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                XSSFCell cell = sheet.getRow(2).getCell(0);
                if (i <= 1) {
                    cell.setCellValue("送货单（幼儿）");
                } else {
                    cell.setCellValue("送货单（老师）");
                }
            }
            workbook.write(new FileOutputStream(file));
        }
    }
}
