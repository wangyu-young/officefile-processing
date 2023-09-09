package fileprocessing.Utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;


public class WordContentHignLightUtil {



    /**
     * 替换word内容
     * @param src 原word文件地址
     * @param dest 转换后word文件地址
     * @param params 替换参数map
     * @throws Exception
     */
    public static Boolean convertWordByXwpf(String src, String dest, Map<String, String> params,Boolean isHighLight) throws Exception {

        Boolean isContains = false;

        String targetDirectory = dest.substring(0, dest.lastIndexOf("\\"));
        File file = new File(targetDirectory);
        if (!file.exists()){
            file.mkdir();
            System.out.println("创建文件夹："+ file.getAbsolutePath());
        }

        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(src));

        // 基础内容替换
        List<XWPFParagraph> xwpfParagraphList = document.getParagraphs();
        for (int i = 0, size = xwpfParagraphList.size(); i < size; i++) {
            XWPFParagraph xwpfParagraph = xwpfParagraphList.get(i);
            for (String key: params.keySet()) {
                String paragraphText = xwpfParagraph.getText();
                if (paragraphText.contains(key)){
                    isContains = true;
                    replaceInParagraph(paragraphText, xwpfParagraph, key, params.get(key),isHighLight);
                }
            }
        }
        // 表格内容替换
        Iterator<XWPFTable> tablesIterator = document.getTablesIterator();
        while (tablesIterator.hasNext()) {
            XWPFTable xwpfTable = tablesIterator.next();
            for (int i = 0, count = xwpfTable.getNumberOfRows(); i < count; i++) {
                XWPFTableRow xwpfTableRow = xwpfTable.getRow(i);
                List<XWPFTableCell> xwpfTableCellList = xwpfTableRow.getTableCells();
                for (int j = 0, cellSize = xwpfTableCellList.size(); j < cellSize; j++) {
                    XWPFTableCell xwpfTableCell = xwpfTableCellList.get(j);
                    List<XWPFParagraph> paragraphList = xwpfTableCell.getParagraphs();
                    for (int k = 0, paragraphSize = paragraphList.size(); k < paragraphSize; k++) {
                        XWPFParagraph xwpfParagraph = paragraphList.get(k);

                        String paragraphText = xwpfParagraph.getText();
                        if (StringUtils.isBlank(paragraphText)) {
                            continue;
                        }
                        for (String key : params.keySet()) {
                            if (paragraphText.contains(key)){
                                isContains = true;
                                replaceInParagraph(paragraphText,xwpfParagraph, key, params.get(key),isHighLight);
                            }
                        }
                    }
                }
            }
        }
        try (FileOutputStream outStream = new FileOutputStream(dest)) {
            document.write(outStream);
        }
        document.close();
        return isContains;
    }

    private static void replaceInParagraph(String paragraphText,XWPFParagraph xwpfParagraph, String oldString, String newString,Boolean isHighLight) {

        List<XWPFRun> runs = xwpfParagraph.getRuns();
        int runSize = runs.size();
        StringBuilder textSb = new StringBuilder();
        Map<Integer, String> textMap = new HashMap<>();

        for (int j = 0; j < runSize; j++) {
            XWPFRun xwpfRun = runs.get(j);
            int textPosition = xwpfRun.getTextPosition();
            String text = xwpfRun.getText(textPosition);
            textSb.append(text);
            textMap.put(j, text);
        }
        // 判断是否重合
        if (!textSb.toString().contains(oldString)){
            return;
        }

        int count = count(textSb.toString(), oldString);

        for (int a = 0; a < count; a++) {

            runs = xwpfParagraph.getRuns();
            runSize = runs.size();

            for (int j = 0; j < runSize; j++) {
                XWPFRun xwpfRun = runs.get(j);
                int textPosition = xwpfRun.getTextPosition();
                String text = xwpfRun.getText(textPosition);
                textMap.put(j, text);
            }

            int startIndex = 0;
            int mapSize = textMap.size();
            int maxEndIndex = oldString.length();
            Integer startPosition = null, endPosition = null;
            String uuid = UUID.randomUUID().toString();
            alwaysFor: for(;;) {
                if (startIndex > mapSize) {
                    break;
                }
                int endIndex = startIndex;
                while (endIndex >= startIndex && maxEndIndex > endIndex - startIndex) {
                    StringBuilder strSb = new StringBuilder();
                    for (int i = startIndex; i <= endIndex; i++) {
                        strSb.append(textMap.getOrDefault(i, uuid));
                    }
                    if (!strSb.toString().trim().equals(oldString)) {
                        ++endIndex;
                    }else {
                        startPosition = startIndex;
                        endPosition = endIndex;
                        break alwaysFor;
                    }
                }
                ++startIndex;
            }

            if (startPosition != null && endPosition != null) {
                XWPFRun modelRun = runs.get(endPosition);
                XWPFRun xwpfRun = xwpfParagraph.insertNewRun(endPosition + 1);
                xwpfRun.setText(newString);
                if (modelRun.getFontSize() != -1) {
                    xwpfRun.setFontSize(modelRun.getFontSize());
                }

                xwpfRun.setFontFamily(modelRun.getFontFamily());
                xwpfRun.setBold(modelRun.isBold());
                if (isHighLight){
                    xwpfRun.setColor("FF3030");
                }else {
                    xwpfRun.setColor(modelRun.getColor());
                }

                for (int i = endPosition; i >= startPosition; i--) {
                    try {
                        xwpfParagraph.removeRun(i);
                    }catch (IllegalArgumentException e){
                        System.out.println(("不支持删除字段或超链接(1)：" + paragraphText));
                        xwpfParagraph.removeRun(endPosition + 1);
                    }
                }
            } else {
                // 最小粒度无法匹配，此处采用下下策粗粒度替换文本
                String text = xwpfParagraph.getText();
                XWPFRun xwpfRun = xwpfParagraph.getRuns().get(0);
                String fontFamily = xwpfRun.getFontFamily();
                int fontSize = xwpfRun.getFontSize();
                XWPFRun insertXwpfRun = null;
                try {
                    insertXwpfRun = xwpfParagraph.insertNewRun(runSize);
                }catch (IndexOutOfBoundsException e){
                    System.out.println("不支持删除字段或超链接(2)："+paragraphText);
                    return;
                }
                if (insertXwpfRun == null){
                    continue;
                }
                insertXwpfRun.setText(text.replace(oldString, newString));
                insertXwpfRun.setFontFamily(fontFamily);
                if (fontSize != -1){
                    insertXwpfRun.setFontSize(fontSize);
                }
                if (isHighLight){
                    insertXwpfRun.setColor("FF3030");
                }else {
                    insertXwpfRun.setColor(xwpfRun.getColor());
                }
                insertXwpfRun.setBold(xwpfRun.isBold());

                for (int i = runSize - 1; i >= 0; i--) {
                    try{
                        xwpfParagraph.removeRun(i);
                    }catch (IllegalArgumentException e){
                        System.out.println("不支持删除字段或超链接(3)："+paragraphText);
                        xwpfParagraph.removeRun(runSize);
                        break;
                    }
                }
            }
        }



    }

    /**
     * str1中str2出现的次数
     * @param str1
     * @param str2
     * @return
     */
    private static int count(String str1,String str2){
        int index = 0;
        int sum = 0;//统计个数
        while (str1.indexOf(str2,index) != -1){//从index的索引处开始查找
            index = str1.indexOf(str2,index) + 1;//加1往下查找
            sum ++;
        }
        return sum;
    }
}
