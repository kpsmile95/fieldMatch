package com.adc;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.TableModel;
import java.io.*;

/**
 * @Author wangmian
 * @Date 2020/10/27
 */
public class ExcelExporter {
    /**导出excel到文件 */
    public static void exportTable(XSSFWorkbook workbook, File file) throws IOException {
        OutputStream os = new FileOutputStream(file);
        if (os == null) {
            return;
        }
        workbook.write(os);
        os.close();
    }
}
