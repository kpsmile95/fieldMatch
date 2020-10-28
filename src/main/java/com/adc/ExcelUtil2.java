package com.adc;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author wangmian
 * @Date 2020/10/27
 */
public class ExcelUtil2 {

    private static Map<Integer, String> errorMap = new HashMap<Integer, String>(){{
        put(0, "Error_chang");
        put(1, "Error_kuan");
        put(2, "Error_gao");
        put(3, "Error_zj");
        put(4, "Error_pl");
        put(5, "Error_gl");
        put(6, "Error_zbzl");
        put(7, "Error_rlzl");
        put(8, "Error_edzk");
    }};

    /* 打开对应的Excel文件 */
    public static XSSFWorkbook validateFile(String fileQCZJ,String fileGG,List<Integer> itemRank) throws IOException {
        FileInputStream fisQCZJ = new FileInputStream(new File(fileQCZJ));
        FileInputStream fisGG = new FileInputStream(new File(fileGG));

        XSSFWorkbook workbookQCZJ = new XSSFWorkbook(fisQCZJ);
        XSSFWorkbook workbookGG = new XSSFWorkbook(fisGG);

        XSSFSheet sheetQCZJ = workbookQCZJ.getSheetAt(0);
        XSSFSheet sheetGG = workbookGG.getSheetAt(0);

        return validate(sheetQCZJ, sheetGG,itemRank);
    }

    private static XSSFWorkbook validate(XSSFSheet sheetQCZJ,XSSFSheet sheetGG,List<Integer> itemRank){
        int qczjNum = sheetQCZJ.getLastRowNum();
        int ggNum = sheetGG.getLastRowNum();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("kx_id");
        row.createCell(1).setCellValue("clxh");
        row.createCell(2).setCellValue("info");
        row.createCell(3).setCellValue("rkrq");

        XSSFRow rowGG = null;
        XSSFRow rowQC = null;

        List<Row> rows = null;

        int rowNum = 1;
        for (int i = 0; i < qczjNum; i++) {
            rowQC = sheetQCZJ.getRow(i + 1);
            String zppKx = rowQC.getCell(1).getStringCellValue();

            rows = new ArrayList<>();

            XSSFRow row1 = sheet.createRow(rowNum);
            row1.createCell(0).setCellValue(rowQC.getCell(0).getStringCellValue());
            row1.createCell(2).setCellValue("");
            row1.createCell(3).setCellValue(LocalDateTime.now().toString());
            rowNum++;

            if (!(zppKx == null || zppKx.isEmpty())) {
                for (int j = 0; j < ggNum; j++) {
                    rowGG = sheetGG.getRow(j + 1);
                    String zpp = rowGG.getCell(1).getStringCellValue();
                    if (!(zpp == null || zpp.isEmpty())) {
                        if (zpp.equals(zppKx)) {
                            rows.add(rowGG);
                            continue;
                        }
                    }

                }
            }
            if (rows.isEmpty()) {
                row1.getCell(2).setCellValue("Error_zpp");
                continue;
            }

            Row finalRow = null;
            List<Integer> matchNum = new ArrayList<>();
            for (Row row2 : rows) {

            }
        }
        return workbook;
    }

    private static boolean validate(int i,String qczj, String gg, Row row){
        boolean flag = false;
        if (!qczj.contains("-") && !gg.contains("-")) {
            String[] qczjs = qczj.split(";");
            for (String qc : qczjs) {
                if (!gg.contains(qc)) {
                    row.getCell(2).setCellValue(errorMap.get(i));
                    flag = true;
                }
            }
        }
        return flag;
    }

    private static String validateByRank(int i, Row rowQC, Row rowGG, Row row) {
        String qczj = null;
        String gg = null;

        boolean flag = false;

        for (int j = 0; j < 9; j++) {
            flag = false;
            if (i == j) {
                qczj = rowQC.getCell(j + 2).getStringCellValue();
                gg = rowGG.getCell(j + 2).getStringCellValue();
                flag = validate(i,qczj, gg, row);
                break;
            }
        }
        String error = "";
        if (flag) {
            error = errorMap.get(i);
        }
        return error;
    }
}
