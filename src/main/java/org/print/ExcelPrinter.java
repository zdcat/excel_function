package org.print;

import java.awt.print.PageFormat;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.standard.MediaPrintableArea;
import javax.print.attribute.standard.MediaSizeName;
import javax.print.attribute.standard.PrintQuality;
import javax.print.attribute.standard.PrinterName;

public class ExcelPrinter {

    public static void main(String[] args) throws IOException, InvalidFormatException, PrinterException {
        // 获得文件路径
        String filePath = "C:\\Users\\84334\\Desktop\\order\\2023\\票\\单子综合\\6.8\\吾悦-兴西-圪僚沟-赞城 加3元\\圪僚沟1.xlsx";


        // 打印机名称
        String printerName = "Jolimark 24-pin printer";
        // 获取打印机
        PrinterName printer = new PrinterName(printerName, null);
        System.out.println(printer);


//
//        // 遍历文件夹
//        File folder = new File(folderPath);
//        if (folder.isDirectory()) {
//            File[] files = folder.listFiles();
//            for (File file : files) {
//                // 判断是否是 Excel 文件
//                if (file.getName().endsWith(".xlsx")) {
//                    // 打开 Excel 文件
//                    Workbook workbook = WorkbookFactory.create(file);
//
//                    // 获取默认的打印设置
//                    PrintSetup printSetup = workbook.getSheetAt(0).getPrintSetup();
//
//                    // 设置打印方向和缩放比例
//                    printSetup.setOrientation(PrintOrientation.LANDSCAPE);
//                    printSetup.setScale((short) 100); // 缩放比例为 100%
//
//                    // 打印所有页
//                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//                        Sheet sheet = workbook.getSheetAt(i);
//                        PrintHelper.printSheet(sheet, printService, attributes);
//                    }
//
//                    // 关闭 Excel 文件
//                    workbook.close();
//                }
//            }
//        }
    }

    private static String getFolderPath() {
        int month = 6;
        int day = 4;

        // 文件夹路径
        String folderPath = "C:\\Users\\84334\\Desktop\\order\\2023\\票\\单子综合\\";
        folderPath = folderPath + month + "." + day;
        System.out.println(folderPath);
        return folderPath;

    }
}
