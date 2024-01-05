package org.test;

import java.io.File;

public class Main {
    public static void main(String[] args) {
//        String root = "C:\\Users\\84334\\Desktop\\order\\2024\\票\\单子综合";
//        File file = new File(root);
//        for (File listFile : file.listFiles()) {
//
//            if (listFile.getName().substring(0, 1).equals("5")) {
//
//                String abso = listFile.getAbsolutePath() + "\\加3元";
//
//                File file1 = new File(abso);
//                for (File listedFile : file1.listFiles()) {
//                    if (listedFile.getName().contains("赞城") && listedFile.getName().contains("补单")){
//                        System.out.println(abso);
//                    }
//                }
//            }
//        }
        // C:\Users\84334\Desktop\order\2024\票\单子综合\5.10\加3元
        // C:\Users\84334\Desktop\order\2024\票\单子综合\5.31\加3元
        Double aDouble = new Double(213.45);
        String chineseString = DoubleToChinese.getChineseString(aDouble);
        System.out.println(chineseString);
    }
}
