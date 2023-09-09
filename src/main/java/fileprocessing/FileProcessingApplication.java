package fileprocessing;


import fileprocessing.service.FileNameReplaceService;
import fileprocessing.service.excel.ExcelContentHignLightService;
import fileprocessing.service.excel.ExcelContentReplaceService;
import fileprocessing.service.word.WordContentHignLightService;
import fileprocessing.service.word.WordContentReplaceService;

import java.util.Scanner;

public class FileProcessingApplication {

    public static final String SOURCE_FILE_PATH = "D:\\files";
    public static final String CONFIG_FILE_PATH = "D:\\文本替换配置表.xlsx";

    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);
        System.out.print("请输入一个整数：\n" +
                "1.Excel文本高亮\n" +
                "2.Excel文本替换\n" +
                "3.Word文本高亮\n" +
                "4.Word文本替换\n" +
                "5.文件名替换\n");
        int number = scanner.nextInt(); // 从标准输入流读取整数

        switch (number){
            case 1:
                ExcelContentHignLightService.ExcelContentHignLight();
                break;
            case 2:
                ExcelContentReplaceService.ExcelContentReplace();
                break;
            case 3:
                WordContentHignLightService.WordContentHignLight();
                break;
            case 4:
                WordContentReplaceService.WordContentReplace();
                break;
            case 5:
                FileNameReplaceService.FileNameReplace();
                break;
        }
    }

}
