package com.moon;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class PoiTest {
    public static void main(String[] args) {
        List<String> headList = new ArrayList<>();
        headList.add("书名");
        headList.add("价格");
        List<String> colNames = new ArrayList<>();
        colNames.add("name");
        colNames.add("price");
        List<MyBook> bodyList = new ArrayList<>();
        bodyList.add(new MyBook("道德经", "100.00"));
        bodyList.add(new MyBook("孙子兵法", "10.00"));
        HSSFWorkbook workbook = PoiDemo.expExcel("书籍管理", headList, colNames, bodyList);
        File desktopDir = FileSystemView.getFileSystemView().getHomeDirectory();
        String desktopPath = desktopDir.getAbsolutePath() + "\\";
        PoiDemo.outFile(workbook, desktopPath + "book.xls");
    }
}
