package org.print;

import javax.print.*;
import java.io.*;

public class PrinterExample {
    public static void main(String[] args) throws IOException, PrintException {

        // 获取系统上所有的打印机
        PrintService[] printServices = PrintServiceLookup.lookupPrintServices(null, null);

        // 选择第一个打印机
        PrintService myPrinter = printServices[0];
        System.out.println(myPrinter.getName());

        // 创建一个打印任务
        DocPrintJob job = myPrinter.createPrintJob();

        // 指定要打印的文件
//        FileInputStream fis = new FileInputStream("file.txt");
//        Doc doc = new SimpleDoc(fis, DocFlavor.INPUT_STREAM.AUTOSENSE, null);
//
//        // 开始打印
//        job.print(doc, null);
    }
}
