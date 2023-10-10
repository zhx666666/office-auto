package io.github.zhx666666.word;

import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

/**
 * 文件工具类
 *
 * @author 明快de玄米61
 * @date 2022/6/28 16:04
 */
public class FileUtil {

    /**
     * 创建临时文件
     * @author 明快de玄米61
     * @date   2022/6/30 16:44
     * @param  fileName 文件名称
     * @return 临时文件
     **/
    public static File createTempFile(String fileName) {
        try {
            // 参数判断
            if (StringUtils.isBlank(fileName) || fileName.lastIndexOf(".") < 0) {
                return null;
            }
            // 创建临时目录
            String tmpdirPath = System.getProperty("java.io.tmpdir");
            String dateStr = new SimpleDateFormat("yyyyMMdd").format(new Date());
            String uuId = UUID.randomUUID().toString().replaceAll("-", "");
            String[] params = {tmpdirPath, dateStr, uuId};
            // 结果例如：C:\Users\Administrator\AppData\Local\Temp\20220630\0f774122c38b423793cc1c121611c142
            String fileUrl = StringUtils.join( params, File.separator);
            if (!new File(fileUrl).exists()) {
                new File(fileUrl).mkdirs();
            }
            // 创建临时文件
            File file = new File(fileUrl, fileName);
            file.createNewFile();
            return file;
        } catch (Exception e) {
            System.out.println("》》》创建临时文件失败，临时文件名称：" + fileName);
            e.printStackTrace();
        }
        return null;
    }



    public static String createTempFilePath(String path,String fileName) {
        try {
            // 参数判断
            if (StringUtils.isBlank(fileName) || fileName.lastIndexOf(".") < 0) {
                return null;
            }
            // 创建临时目录
            //String tmpdirPath = System.getProperty("java.io.tmpdir");
            String dateStr = new SimpleDateFormat("yyyyMMdd").format(new Date());
            String dateStr2 = new SimpleDateFormat("HH:mm:ss").format(new Date());
            fileName = dateStr2+"_"+fileName;
            //String uuId = UUID.randomUUID().toString().replaceAll("-", "");
            String[] params = {path, dateStr};
            // 结果例如：C:\Users\Administrator\AppData\Local\Temp\20220630\0f774122c38b423793cc1c121611c142
            String fileUrl = StringUtils.join( params, File.separator);
            if (!new File(fileUrl).exists()) {
                new File(fileUrl).mkdirs();
            }
            // 创建临时文件
            File file = new File(fileUrl, fileName);
            file.createNewFile();
            return fileUrl+"/"+fileName;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 删除父级文件 / 目录
     * @param sourceFiles 当前文件/目录
     */
    public static void deleteParentFile(File... sourceFiles) {
        if (sourceFiles != null && sourceFiles.length > 0) {
            // 查找父级目录
            List<File> parentFiles = new ArrayList<>(sourceFiles.length);
            for (File sourceFile : sourceFiles) {
                if (sourceFile != null && sourceFile.exists()) {
                    parentFiles.add(sourceFile.getParentFile());
                }
            }
            // 删除父级目录
            deleteFile(parentFiles.toArray(new File[0]));
        }
    }

    /**
     * 删除文件 / 目录
     * @author 明快de玄米61
     * @date   2022/6/28 16:04
     * @param  sourceFiles 当前文件/目录
     **/
    public static void deleteFile(File... sourceFiles) {
        if (sourceFiles != null && sourceFiles.length > 0) {
            for (File sourceFile : sourceFiles) {
                try {
                    // 判断存在性
                    if (sourceFile == null || !sourceFile.exists()) {
                        continue;
                    }
                    // 判断文件类型
                    if (sourceFile.isDirectory()) {
                        // 遍历子级文件 / 目录
                        File[] childrenFile = sourceFile.listFiles();
                        if (childrenFile != null && childrenFile.length > 0) {
                            for (File childFile : childrenFile) {
                                // 删除子级文件 / 目录
                                deleteFile(childFile);
                            }
                        }
                    }
                    // 删除 文件 / 目录 本身
                    sourceFile.delete();
                } catch (Exception e) {
                    System.out.println("》》》删除文件/目录报错，其中文件/目录全路径：" + sourceFile.getAbsolutePath());
                    e.printStackTrace();
                }
            }
        }
    }

}

