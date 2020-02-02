package ToolsFiles;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;

import static ToolsFiles.library.StoreUtil.*;
import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreWriting.*;

public class TFMagnetsReport {

    public static void main(String directory, String reportFileName) {

        //grabReportsMagnetInOneFile(getWorkDir() + directory, getWorkDir() + directory + reportFileName);

        HSSFWorkbook book = createBook(getWorkDir() + directory + "tf report.xls");
        HashMap map = createMapOrderSingle(book, "июнь.xls",0,0,1);
        closeBook(book);




        HSSFWorkbook book0 = createBook("C:/Java/Projects/JavaTest/CommonSalesReport/commonReport.xls");
        ArrayList listArts = createVertListOfValues(book0, "Чистый бланк", 2,0);
        closeBook(book0);

        ArrayList ar = createLongListOrderFromMultipleMap(map, createListOfArtsFromPassport());


        HSSFWorkbook book2 = createBook(getWorkDir() + directory + "tf report.xls");

        HSSFSheet sheet = book2.createSheet("OOO");

        for (int i = 0; i < ar.size(); i++) {
            ArrayList l = (ArrayList)ar.get(i);
            HSSFRow row = sheet.createRow(i);
            HSSFCell artCell = row.createCell(0);
            HSSFCell amCell = row.createCell(1);
            artCell.setCellValue((String)l.get(0));
            System.out.println((String)l.get(0));

            if (l.get(1) == null) {continue;}
            amCell.setCellValue((Integer)l.get(1));
        }

        writeBook(book2, getWorkDir() + directory + "tf report.xls" );
        closeBook(book2);

        //openFile(getWorkDir() + reportFileName);
    }

    // Take all TF reports from directory and put it in report file
    public static void grabReportsMagnetInOneFile(String directoryName, String reportFileName) {

        ArrayList<String> directoryList = createDirectoryList(directoryName);

        for (String file : directoryList) {

            String fileName = directoryName + file;

            System.out.println(file);

            HSSFWorkbook book = createBook(fileName);
            ArrayList orderList = createListOrderShort(book, book.getSheetAt(0).getSheetName(),7,0,5);
            //printArray(orderList);
            closeBook(book);

            HSSFWorkbook repBook = createBook(reportFileName);
            writeVertListInBookSingle(repBook,file,0,0,1,orderList);
            writeBook(repBook, reportFileName);
            closeBook(repBook);

        }

    }
}
