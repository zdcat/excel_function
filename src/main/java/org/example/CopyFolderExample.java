package org.example;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

/**
 * 每次都先把原来所有的加价的单子都删了，然后重新创建加价的单子
 */
public class CopyFolderExample {
    public static void main(String[] args) throws IOException {
        String userName = System.getProperty("user.name");
        File sourceFolder = new File("C:\\Users\\" + userName + "\\Desktop\\order\\2023\\票\\单子综合\\" + args[0] + "\\正常价格");
        File destinationFolder = new File("C:\\Users\\" + userName +
                "\\Desktop\\order\\2023\\票\\单子综合\\" + args[0] + "\\" + getFileName(args) + " 加3元");

        // 如果目标文件夹已存在，则删除它及其内容
        if (destinationFolder.exists()) {
            deleteFolder(destinationFolder);
        }


        // 复制前删除指定日期目录下所有包含“加”的目录，以删除之前的部分加价情况和现在加价情况不一样的文件
        // 比如原来加3家，那么现在加4家的话按照上面的逻辑是删不掉原来的加3家的，所以需要找到旧的加价的文件夹，然后删除
        deleteOldAddFolder(args, userName);


        // 创建目标文件夹
        destinationFolder.mkdir();

        // 获取源文件夹中的所有文件和子文件夹
        File[] files = sourceFolder.listFiles();

        // 遍历所有文件，递归复制子文件夹
        for (File file : files) {
            if (file.isDirectory()) {
                String sourceFolderPath = file.getAbsolutePath();
                String destinationFolderPath = destinationFolder.getAbsolutePath() + File.separator + file.getName();
                copyFolder(new File(sourceFolderPath), new File(destinationFolderPath));
            } else {
                File destinationFile = new File(destinationFolder.getAbsolutePath() + File.separator + file.getName());
                Files.copy(file.toPath(), destinationFile.toPath());
            }
        }
    }

    private static void deleteOldAddFolder(String[] args, String userName) throws IOException {
        File folder = new File("C:\\Users\\" + userName +
                "\\Desktop\\order\\2023\\票\\单子综合\\" + args[0]);
        File[] files = folder.listFiles();
        // 遍历所有文件，递归删除子文件夹
        for (File file : files) {
            if (file.isDirectory()) {
                if (file.getName().contains("加"))
                    deleteFolder(file);
            } else {
                file.delete();
            }
        }

        // 删除空文件夹
        folder.delete();
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


    private static void copyFolder(File sourceFolder, File destinationFolder) throws IOException {
        // 创建目标文件夹
        destinationFolder.mkdir();

        // 获取源文件夹中的所有文件和子文件夹
        File[] files = sourceFolder.listFiles();

        // 遍历所有文件，递归复制子文件夹
        for (File file : files) {
            if (file.isDirectory()) {
                String sourceFolderPath = file.getAbsolutePath();
                String destinationFolderPath = destinationFolder.getAbsolutePath() + File.separator + file.getName();
                copyFolder(new File(sourceFolderPath), new File(destinationFolderPath));
            } else {
                File destinationFile = new File(destinationFolder.getAbsolutePath() + File.separator + file.getName());
                Files.copy(file.toPath(), destinationFile.toPath());
            }
        }
    }

    private static void deleteFolder(File folder) throws IOException {
        // 获取文件夹中的所有文件和子文件夹
        File[] files = folder.listFiles();

        // 遍历所有文件，递归删除子文件夹
        for (File file : files) {
            if (file.isDirectory()) {
                deleteFolder(file);
            } else {
                file.delete();
            }
        }

        // 删除空文件夹
        folder.delete();
    }
}
