package com.hidata.tool;

import java.io.File;
import java.io.IOException;

/**
 * 创建新文件和目录
 */
public class ZhongYinCreateFileUtil {

    /**
     * 创建目录
     * @param destDirName   目标目录名
     * @return 目录创建成功放回true，否则返回false
     */
    public static boolean createDir(String destDirName) {
        File dir = new File(destDirName);
        if (dir.exists()) {
            System.out.println("创建目录" + destDirName + "失败，目标目录已存在！");
            return false;
        }
        if (!destDirName.endsWith(File.separator)) {
            destDirName = destDirName + File.separator;
        }
        // 创建目标目录
        if (dir.mkdir()) {
            System.out.println("创建目录" + destDirName + "成功！");
            return true;
        } else {
            System.out.println("创建目录" + destDirName + "失败！");
            return false;
        }
    }

    /**
     * 创建临时文件
     * @param prefix    临时文件名的前缀
     * @param suffix    临时文件名的后缀
     * @param dirName   临时文件所在的目录，如果输入null，则在用户的文档目录下创建临时文件
     * @return 临时文件创建成功返回true，否则返回false
     */
    public static String createTempFile(String prefix, String suffix,
                                        String dirName) {
        File tempFile = null;
        if (dirName == null) {
            try {
                // 在默认文件夹下创建临时文件
                tempFile = File.createTempFile(prefix, suffix);
                // 运行结束是删除
                tempFile.deleteOnExit();
                // 立即删除
                tempFile.delete();
                // 返回临时文件的路径
                return tempFile.getCanonicalPath();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("创建临时文件失败!" + e.getMessage());
                return null;
            }
        } else {
            File dir = new File(dirName);
            // 如果临时文件所在目录不存在，首先创建
            if (!dir.exists()) {
                if (ZhongYinCreateFileUtil.createDir(dirName)) {
                    System.out.println("创建临时文件失败，不能创建临时文件所在的目录！");
                    return null;
                }
            }
            try {
                // 在指定目录下创建临时文件
                tempFile = File.createTempFile(prefix, suffix, dir);
                return tempFile.getCanonicalPath();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("创建临时文件失败!" + e.getMessage());
                return null;
            }
        }
    }

    public static void main(String[] args) {
        // 创建临时文件
        String prefix = "temp";
        String surfix = ".txt";
        for (int i = 0; i < 10; i++) {
            System.out.println("创建了临时文件: "
                    + ZhongYinCreateFileUtil.createTempFile(prefix, surfix, null));
        }
    }
}
