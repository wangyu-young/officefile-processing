package fileprocessing.service;


import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class FileNameReplaceService {

    public static void FileNameReplace(){
        try {
            File file = new File("D:\\files");
            //处理文件
            Map<String, String> map = loadConfig();
            recursiveTraversalFolder(file,map);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Map<String,String> loadConfig() throws IOException {
        String intPath = "D:\\文本替换配置表.xlsx";

        File file = new File(intPath);
        if (!file.exists()){
            throw new RuntimeException("D:\\文本替换配置表.xlsx (系统找不到指定的文件。)");
        }

        FileInputStream fs = new FileInputStream(intPath);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheet("文件名替换");

        //遍历行
        Map<String,String> map = new HashMap<>();

        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {

            XSSFRow row = sheet.getRow(i);
            String key = "";
            XSSFCell cell = row.getCell(0);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                cell.setCellType(CellType.STRING);
            }else {
                continue;
            }
            key = cell.getStringCellValue();

            String value = "";
            XSSFCell cell1 = row.getCell(1);
            if (cell1 != null && cell.getCellType() != CellType.BLANK) {
                cell1.setCellType(CellType.STRING);
            }else {
                continue;
            }
            value = cell1.getStringCellValue();

            map.put(key,value);
        }
        return map;
    }

    public static void recursiveTraversalFolder(File folder,Map<String, String> map) {

        if (folder.exists()) {
            File[] fileArr = folder.listFiles();
            if (null == fileArr || fileArr.length == 0) {
                System.out.println("文件夹是空的!");
                return;
            } else {
                File newDir = null;//文件所在文件夹路径+新文件名
                String newName = "";//新文件名
                String fileName = null;//旧文件名
                File parentPath = new File("");//文件所在父级路径
                for (File file : fileArr) {
                    if (file.isDirectory()) {//是文件夹，继续递归，如果需要重命名文件夹，这里可以做处理
                        System.out.println("文件夹:" + file.getAbsolutePath() + "，继续递归！");
                        recursiveTraversalFolder(file, map);
                    } else {//是文件，判断是否需要重命名
                        fileName = file.getName();
                        parentPath = file.getParentFile();

                        for (String key : map.keySet()){
                            if (fileName.contains(key)) {//文件名包含需要被替换的字符串
                                newName = fileName.replaceAll(key, map.get(key));//新名字
                                newDir = new File(parentPath + "/" + newName);//文件所在文件夹路径+新文件名
                                file.renameTo(newDir);//重命名
                                System.out.println("修改后：" + newDir);
                            }
                        }
                    }
                }
            }
        } else {
            System.out.println("文件不存在!");
        }
    }
}
