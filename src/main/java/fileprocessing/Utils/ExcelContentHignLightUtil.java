package fileprocessing.Utils;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;

public class ExcelContentHignLightUtil {

    /**
     * 替换Excel模板文件内容
     *
     * @param list     需要替换的标签建筑队形式
     * @param intPath Excel模板文件路径
     * @param outPath Excel生成文件路径
     */
    public static Boolean replaceSheetsModel(String intPath, String outPath, List<String> list) {

        Boolean isContains = false;

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

            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


            for (int j = 0; j < workbook.getNumberOfSheets(); j++) {
                sheet = workbook.getSheetAt(j);
                Iterator rows = sheet.rowIterator();
                while (rows.hasNext()) {
                    XSSFRow row = (XSSFRow) rows.next();
                    if (row != null) {
                        int num = row.getLastCellNum();
                        for (int i = 0; i < num; i++) {
                            XSSFCell cell = row.getCell(i);
                            if (cell == null || cell.getCellType() == CellType.BLANK) {
                                continue;
                            }
                            if (cell.getCellType() == CellType.STRING){
                                String value = cell.getStringCellValue();
                                if (!"".equals(value)) {
                                    Iterator<String> it = list.iterator();
                                    while (it.hasNext()) {
                                        String text = it.next();
                                        if (value.contains(text)) {
                                            cell.setCellStyle(cellStyle);
                                            isContains = true;
                                            break;
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
        return isContains;
    }
}
