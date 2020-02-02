package ToolsFiles.java;

import static ToolsFiles.library.StoreUtil.*;
import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreWriting.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;

public class CommonSalesReport {

    static String workDir = "C:/Java/Projects/JavaTest/CommonSalesReport/";
    static String reportFile = workDir + "commonReport.xls";

    public static void main(String[] args) {

        //reportFromOldReportToNew("зарядье.xls", "report", "Зарядье");
        //reportFromOldReportToNew("дом книги.xls", "report", "Дом Книги");
        //reportFromOldReportToNew("ип афанасьев.xls", "report", "Бон Аппетит");
        //reportFromOldReportToNew("норма 8.xls", "report", "Норма 8");
        //reportFromOldReportToNew("rawReport.xls", "Бука", "Бука");
        //reportFromOldReportToNew("rawReport.xls", "Буквоед", "Буквоед");
        //reportFromOldReportToNew("rawReport.xls", "Пароль НН", "Пароль НН");

        openFile(reportFile);

    }

    public static void reportFromOldReportToNew(String fileName, String oldSheetName, String newSheetName) {

        String file = workDir + fileName;

        HSSFWorkbook book = createBook(file);
        HashMap<String, ArrayList> map = createMapOrderMultiple(book, oldSheetName, 3, 0, 2, 1);
        closeBook(book);


        HSSFWorkbook repBook = createBook(reportFile);

        ArrayList<String> artList = createVertListOfValues(repBook, newSheetName, 2, 0);
        ArrayList<ArrayList> orderList = createLongListOrderFromMultipleMap(map, artList);

        System.out.println("Осталось после перевода в List:");
        printHashMap(map);

        writeVertListInBookMultiple(repBook, newSheetName, 2,0,2, orderList);

        writeBook(repBook, reportFile);
        closeBook(repBook);
    }

}
