package com.adc;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @Author wangmian
 * @Date 2020/10/27
 */
public class ExcelUtil2 {

    //报错类型
    private static Map<Integer, String> errorMap = new HashMap<Integer, String>() {{
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

    //对数据进行校验
    public static XSSFWorkbook validateFile(String fileQCZJ, String fileGG, List<Integer> itemRank) throws IOException {
        FileInputStream fisQCZJ = new FileInputStream(new File(fileQCZJ));
        FileInputStream fisGG = new FileInputStream(new File(fileGG));

        XSSFWorkbook workbookQCZJ = new XSSFWorkbook(fisQCZJ);
        XSSFWorkbook workbookGG = new XSSFWorkbook(fisGG);

        XSSFSheet sheetQCZJ = workbookQCZJ.getSheetAt(0);
        XSSFSheet sheetGG = workbookGG.getSheetAt(0);

        return validate(sheetQCZJ, sheetGG, itemRank);
    }

    private static XSSFWorkbook validate(XSSFSheet sheetQCZJ, XSSFSheet sheetGG, List<Integer> itemRank) {
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

            //获取当前id对应得所有款型
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
            //若是当前id对应得款型没有数据则报错
            if (rows.isEmpty()) {
                XSSFRow row1 = sheet.createRow(rowNum);
                row1.createCell(0).setCellValue(rowQC.getCell(0).getStringCellValue());
                row1.createCell(2).setCellValue("");
                row1.createCell(3).setCellValue(LocalDateTime.now().toString());
                rowNum++;
                row1.getCell(2).setCellValue("Error_zpp");
                continue;
            }
            //循环匹配，检验每条数据，记录每条数据匹配到得最大的列
            Row finalRow = null;
            List<Integer> matchNum = new ArrayList<>();
            for (int j = 0; j < rows.size(); j++) {
                Row cells = rows.get(j);
                for (int k = 0; k < itemRank.size(); k++) {
                    Integer integer = itemRank.get(k);
                    if (!validateByRank(integer, rowQC, cells)) {
                        matchNum.add(integer);
                        break;
                    }
                    if (k == 8) {
                        matchNum.add(9);
                    }
                }
            }
            Integer max = Collections.max(matchNum);
            int index = matchNum.indexOf(max);
            finalRow = rows.get(index);
            if (max == 9) {
                for (int j = 0; j < matchNum.size(); j++) {
                    if (matchNum.get(j) == 9) {
                        finalRow = rows.get(j);
                        XSSFRow row1 = sheet.createRow(rowNum);
                        row1.createCell(0).setCellValue(rowQC.getCell(0).getStringCellValue());
                        row1.createCell(2).setCellValue("");
                        row1.createCell(3).setCellValue(LocalDateTime.now().toString());
                        rowNum++;
                        row1.createCell(1).setCellValue(finalRow.getCell(0).getStringCellValue());
                    }
                }
            } else {
                XSSFRow row1 = sheet.createRow(rowNum);
                row1.createCell(0).setCellValue(rowQC.getCell(0).getStringCellValue());
                row1.createCell(2).setCellValue("");
                row1.createCell(3).setCellValue(LocalDateTime.now().toString());
                rowNum++;
                row1.getCell(2).setCellValue(errorMap.get(max));
            }
        }
        return workbook;
    }

    private static boolean validate(int i, String qczj, String gg) {
        boolean flag = false;
        if (!qczj.equals("-") && !gg.equals("-")) {
            String[] qczjs = qczj.split(";");
            String[] ggs = gg.split(";");
            for (String qc : qczjs) {
                for (String g : ggs) {
                    if (qc.equals(g)) {
                        flag = true;
                        break;
                    }
                }
                if (flag) {
                    break;
                }
            }
        }
        if (qczj.equals("-") || gg.equals("-")) {
            flag = true;
        }
        return flag;
    }

    private static boolean validateByRank(int i, Row rowQC, Row rowGG) {
        String qczj = null;
        String gg = null;

        boolean validate = false;

        for (int j = 0; j < 9; j++) {
            if (i == j) {
                qczj = rowQC.getCell(j + 2).getStringCellValue();
                gg = rowGG.getCell(j + 2).getStringCellValue();
                validate = validate(i, qczj, gg);
                break;
            }
        }
        return validate;
    }
}
