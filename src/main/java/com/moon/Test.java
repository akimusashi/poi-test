package com.moon;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONArray;

public class Test {
    public static HSSFWorkbook expExcel(JSONArray head, JSONArray body) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("书籍信息");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = null;
        for (int i = 0, len = head.length(); i < len; i++) {
            cell = row.createCell(i);
            cell.setCellValue(head.getString(i));
        }
        for (int i = 0, len = body.length(); i < len; i++) {

        }
        return null;
    }
}
