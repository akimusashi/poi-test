package com.moon;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.List;

public class PoiDemo {
    public static <E> HSSFWorkbook expExcel(String sheetName, List<String> headList, List<String> colNames, List<E> bodyList) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = null;
        for (int i = 0, size = headList.size(); i < size; i++) {
            cell = row.createCell(i);
            cell.setCellValue(headList.get(i));
        }
        /*JSONArray names = new JSONArray(colNames);
        for (int i = 0, size = bodyList.size(); i < size; i++) {
            row = sheet.createRow(i + 1);
            JSONArray rowInfo = new JSONObject(bodyList.get(i)).toJSONArray(names);
            for (int j = 0, len = rowInfo.length(); j < len; j++) {
                cell = row.createCell(j);
                cell.setCellValue(rowInfo.getString(j));
            }
        }*/
        JSONArray names = JSONArray.fromObject(colNames);
        for (int i = 0, size = bodyList.size(); i < size; i++) {
            row = sheet.createRow(i + 1);
            JSONArray rowInfo = JSONObject.fromObject(bodyList.get(i)).toJSONArray(names);
            for (int j = 0, row_size = rowInfo.size(); j < row_size; j++) {
                cell = row.createCell(j);
                cell.setCellValue(rowInfo.getString(j));
            }
        }
        return workbook;
    }

    public static void outFile(HSSFWorkbook workbook, String path) {
        OutputStream os = null;
        try {
            os = new FileOutputStream(path);
            workbook.write(os);
        } catch (FileNotFoundException e) {
            System.out.println("------FileNotFoundException------");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("------IOException------");
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    System.out.println("------IOException------");
                    e.printStackTrace();
                }
            }
        }
    }
}
