package fileprocessing.service.excel;

import fileprocessing.Utils.ExcelContentHignLightUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static fileprocessing.FileProcessingApplication.CONFIG_FILE_PATH;
import static fileprocessing.FileProcessingApplication.SOURCE_FILE_PATH;

public class ExcelContentHignLightService {

    public static void ExcelContentHignLight(){
        try {
            File file = new File(SOURCE_FILE_PATH);

            if (!file.exists()){
                throw new RuntimeException("系统找不到指定的文件："+SOURCE_FILE_PATH);
            }

            //处理文件
            List<String> list = loadConfig();

            List<String> fileNameList = new ArrayList<>();
            listFilesForFolder(file,list,fileNameList);

            System.out.println("含有高亮文本的文件："+fileNameList);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<String> loadConfig() throws IOException {

        File file = new File(CONFIG_FILE_PATH);
        if (!file.exists()){
            throw new RuntimeException("系统找不到指定的文件");
        }

        FileInputStream fs = new FileInputStream(CONFIG_FILE_PATH);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheet("excel文本高亮");

        List<String> list = new ArrayList<>();

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

            list.add(key);
        }
        return list;
    }

    private static void listFilesForFolder(File folder,List<String> list,List<String> fileNameList) throws Exception {

        for (File file : folder.listFiles()) {
            if (file.isDirectory()) {
                listFilesForFolder(file,list,fileNameList);
            } else {
                String fileExtension = getFileExtension(file);
                if (fileExtension.equals("xlsx")) {
                    System.out.println(file.getAbsolutePath());
                    Boolean isContains = ExcelContentHignLightUtil.replaceSheetsModel(file.getAbsolutePath(),
                            file.getAbsolutePath().replace("files", "newfiles"), list);
                    if (isContains){
                        fileNameList.add(file.getName());
                    }
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
