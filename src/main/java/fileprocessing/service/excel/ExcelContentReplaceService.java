package fileprocessing.service.excel;

import fileprocessing.Utils.ExcelContentReplaceUtil;
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

import static fileprocessing.FileProcessingApplication.CONFIG_FILE_PATH;
import static fileprocessing.FileProcessingApplication.SOURCE_FILE_PATH;

/**
 * 替换excel中的文本
 */
public class ExcelContentReplaceService {


    public static void ExcelContentReplace(){
        try {
            File file = new File(SOURCE_FILE_PATH);

            if (!file.exists()){
                throw new RuntimeException("系统找不到指定的文件");
            }
            //处理文件
            Map<String, String> map = loadConfig();

            listFilesForFolder(file,map);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private static Map<String,String> loadConfig() throws IOException {

        File file = new File(CONFIG_FILE_PATH);
        if (!file.exists()){
            throw new RuntimeException("系统找不到指定的配置文件");
        }

        FileInputStream fs = new FileInputStream(CONFIG_FILE_PATH);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheet("excel文本替换");

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

        workbook.close();
        return map;
    }

    private static void listFilesForFolder(File folder,Map<String, String> map) throws Exception {

        for (File file : folder.listFiles()) {
            if (file.isDirectory()) {
                listFilesForFolder(file,map);
            } else {
                String fileExtension = getFileExtension(file);
                if (fileExtension.equals("xls") || fileExtension.equals("xlsx")){
                    System.out.println(file.getAbsolutePath());

                    ExcelContentReplaceUtil.replaceSheetsModel(file.getAbsolutePath(),
                            file.getAbsolutePath().replace("files", "newfiles"), map);
                }
            }
        }
    }

    private static String getFileExtension(File file) {
        String fileName = file.getName();
        if(fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
            return fileName.substring(fileName.lastIndexOf(".")+1);
        else return "";
    }
}
