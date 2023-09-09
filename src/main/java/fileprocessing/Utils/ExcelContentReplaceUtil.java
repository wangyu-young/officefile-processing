package fileprocessing.Utils;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

public class ExcelContentReplaceUtil {

    /**
     * 替换Excel模板文件内容
     *
     * @param map     需要替换的标签建筑队形式
     * @param intPath Excel模板文件路径
     * @param outPath Excel生成文件路径
     */
    public static void replaceSheetsModel(String intPath, String outPath, Map<String,String> map) {
        try {
            String targetDirectory = outPath.substring(0, outPath.lastIndexOf("\\"));
            File file = new File(targetDirectory);
            if (!file.exists()){
                file.mkdir();
                System.out.println("创建文件夹："+ file.getAbsolutePath());
            }

            FileInputStream fs = new FileInputStream(intPath);
            XSSFWorkbook workbook = new XSSFWorkbook(fs);
            XSSFSheet sheet;
            for (int j = 0; j < workbook.getNumberOfSheets(); j++) {
                sheet = workbook.getSheetAt(j);
                Iterator rows = sheet.rowIterator();
                while (rows.hasNext()) {
                    XSSFRow row = (XSSFRow) rows.next();
                    if (row != null) {
                        int num = row.getLastCellNum();
                        for (int i = 0; i < num; i++) {
                            XSSFCell cell = row.getCell(i);
//                            if (cell != null) {
//                                cell.setCellType(CellType.STRING);
//                            }
                            if (cell == null || cell.getCellType() == CellType.BLANK) {
                                continue;
                            }
                            if (cell.getCellType() == CellType.STRING){
                                String value = cell.getStringCellValue();
                                if (!"".equals(value)) {
                                    Set<String> keySet = map.keySet();
                                    Iterator<String> it = keySet.iterator();
                                    while (it.hasNext()) {
                                        String text = it.next();
                                        if (value.contains(text)) {
                                            String newValue = value.replace(text,map.get(text));
                                            value = newValue;
                                            cell.setCellValue(newValue);
                                        }
                                    }
                                } else {
                                    cell.setCellValue("");
                                }
                            }
                        }
                    }
                }
            }
            FileOutputStream fileOut = new FileOutputStream(outPath);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {

            e.printStackTrace();
        }

    }
}
