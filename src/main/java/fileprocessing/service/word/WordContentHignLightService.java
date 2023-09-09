package fileprocessing.service.word;

import fileprocessing.Utils.WordContentHignLightUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import static fileprocessing.FileProcessingApplication.CONFIG_FILE_PATH;
import static fileprocessing.FileProcessingApplication.SOURCE_FILE_PATH;

/**
 * 高量word中的文本
 */
public class WordContentHignLightService {


    public static void WordContentHignLight(){
        try {
            File file = new File(SOURCE_FILE_PATH);
            //处理文件
            if (!file.exists()){
                throw new RuntimeException("系统找不到指定的文件");
            }
            Map<String, String> map = loadConfig();

            List<String> fileNameList = new ArrayList<>();
            listFilesForFolder(file,map,fileNameList);

            System.out.println("含有高亮文本的文件："+fileNameList);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String,String> loadConfig() throws IOException {

        File file = new File(CONFIG_FILE_PATH);
        if (!file.exists()){
            throw new RuntimeException("系统找不到指定的文件");
        }

        FileInputStream fs = new FileInputStream(CONFIG_FILE_PATH);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheet("word文本高亮");

        //遍历行
        Map<String, String> map = new HashMap<>();

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

            map.put(key,key);
        }
        return map;
    }

    private static void listFilesForFolder(File folder,Map<String, String> map,List<String> fileNameList) throws Exception {

        for (File file : folder.listFiles()) {
            if (file.isDirectory()) {
                listFilesForFolder(file,map,fileNameList);
            } else {
                String fileExtension = getFileExtension(file);
                if (fileExtension.equals("docx")) {
                    System.out.println(file.getAbsolutePath());
                    Boolean isContains = WordContentHignLightUtil.convertWordByXwpf(file.getAbsolutePath(),
                            file.getAbsolutePath().replace("files", "newfiles"), map, true);
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
