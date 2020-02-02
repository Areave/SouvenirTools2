package ToolsFiles.library;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;

import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreUtil.*;

public class StoreWriting {

    //////////////////////////////////////////// WRITING


    //Write ArrayList of values of multiple types in particular existing book
    public static void writeVertListInBookSingle(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, ArrayList list) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            sheet = book.createSheet(sheetName);
        }

        for (int i = firstRowNum; i < list.size() + firstRowNum; i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                row = sheet.createRow(i);
            }

            HSSFCell cell = row.createCell(columnNum);

            if (list.get(i) instanceof String) {
                // System.out.println("String, i = " + i + ", val " + prices.get(i));
                cell.setCellValue((String) list.get(i));
            }

            if (list.get(i) instanceof Long) {
                // System.out.println("Long, i = " + i + ", val " + prices.get(i));
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue((Long) list.get(i));
            }

            if (list.get(i) instanceof Integer) {
                //  System.out.println("Integer, i = " + i + ", val " + prices.get(i));
                cell.setCellValue((Integer) list.get(i));
            }

            if (list.get(i) instanceof Double) {
                //  System.out.println("Double, i = " + i + ", val " + prices.get(i));
                cell.setCellValue((Double) list.get(i));
            }

            if (list.get(i) == null) {
                //  System.out.println("Double, i = " + i + ", val " + prices.get(i));
                cell.setCellValue("");
            }
        }

    }



    // IN ArrayList <String or anything>
    public static void writeStringArrayList(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, ArrayList<String> list) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            sheet = book.createSheet(sheetName);
        }

        for (int i = firstRowNum; i < list.size() + firstRowNum; i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                row = sheet.createRow(i);
            }
            HSSFCell cell = row.createCell(columnNum);

            cell.setCellValue(list.get(i));


            if (list.get(i) == null) {
                cell.setCellValue("");
            }
        }

    }


    // Object o = you can put 0, null or "",
    // if 0 - empty values will be 0
    // if null 0 empty values will be skipped
    //  if "" - empty values will be skipped
    //

    // IN ArrayList <ArrayList<art, am>>
    //
    //
    public static void writeArrayArrayStringInt(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, int amNum, ArrayList<ArrayList> list, Object o) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            sheet = book.createSheet(sheetName);
        }

        for (int i = 0; i < list.size(); i++) {

            HSSFRow row = sheet.getRow(i + firstRowNum);
            if (row == null) {
                row = sheet.createRow(i + firstRowNum);
            }

            ArrayList l = list.get(i);

            HSSFCell artCell = row.createCell(columnNum);
            HSSFCell amCell = row.createCell(amNum);

            if ((l.get(1) == null) || ((Integer) l.get(1) == 0)) {

                if (o == null) {
                    firstRowNum--;
                    continue;
                }

                if (o.toString().equals("")) {
                    artCell.setCellValue(l.get(0).toString());
                    continue;
                }
                if ((Integer) o == 0) {
                    artCell.setCellValue(l.get(0).toString());
                    amCell.setCellValue(0);
                    continue;
                }
            }

            artCell.setCellValue(l.get(0).toString());
            amCell.setCellValue((Integer) l.get(1));
        }
    }

    // IN ArrayList <ArrayList<art, name, am>>
    //
    //
    public static void writeArrayArrayStringNameInt(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, int nameNum, int amNum, ArrayList<ArrayList> list, Object o) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            sheet = book.createSheet(sheetName);
        }

        for (int i = 0; i < list.size(); i++) {

            HSSFRow row = sheet.getRow(i + firstRowNum);
            if (row == null) {
                row = sheet.createRow(i + firstRowNum);
            }

            ArrayList l = list.get(i);

            HSSFCell artCell = row.createCell(columnNum);
            HSSFCell nameCell = row.createCell(nameNum);
            HSSFCell amCell = row.createCell(amNum);

            if ((l.get(2) == null) || ((Integer) l.get(2) == 0)) {

                if (o == null) {
                    firstRowNum--;
                    continue;
                }

                if (o.toString().equals("")) {
                    artCell.setCellValue(l.get(0).toString());
                    nameCell.setCellValue(l.get(1).toString());
                    continue;
                }
                if ((Integer) o == 0) {
                    artCell.setCellValue(l.get(0).toString());
                    nameCell.setCellValue(l.get(1).toString());
                    amCell.setCellValue(0);
                    continue;
                }
            }

            artCell.setCellValue(l.get(0).toString());
            nameCell.setCellValue(l.get(1).toString());
            amCell.setCellValue((Integer) l.get(2));
        }
    }

    // IN ArrayList <ArrayList<art, Array<am>>>
    //
    //
    public static void writeArrayArrayStringArray(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, int firstAmNum, ArrayList<ArrayList> list, Object o) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            sheet = book.createSheet(sheetName);
        }

        for (int i = 0; i < list.size(); i++) {

            HSSFRow row = sheet.getRow(i + firstRowNum);
            if (row == null) {
                row = sheet.createRow(i + firstRowNum);
            }

            ArrayList l = list.get(i);
            ArrayList am = (ArrayList) l.get(1);

            HSSFCell artCell = row.createCell(columnNum);

            if (am == null) {

                if (o == null) {
                    firstRowNum--;
                    continue;
                }

                if (o.toString().equals("")) {
                    artCell.setCellValue(l.get(0).toString());
                    continue;
                }

                if ((Integer) o == 0) {
                    artCell.setCellValue(l.get(0).toString());
                    row.createCell(firstAmNum).setCellValue(0);
                    continue;
                }
            }

            artCell.setCellValue(l.get(0).toString());
            for (int j = 0; j < am.size(); j++) {

                HSSFCell amCell = row.createCell(firstAmNum + j);
                amCell.setCellValue((Integer) am.get(j));

            }
        }

    }

    // IN ArrayList <ArrayList<art, name, Array<am>>>
    //
    //
    public static void writeArrayArrayStringNameArray(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, int nameNum, int firstAmNum, ArrayList<ArrayList> list, Object o) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            sheet = book.createSheet(sheetName);
        }

        for (int i = 0; i < list.size(); i++) {

            HSSFRow row = sheet.getRow(i + firstRowNum);
            if (row == null) {
                row = sheet.createRow(i + firstRowNum);
            }

            ArrayList l = list.get(i);
            ArrayList am = (ArrayList) l.get(2);

            HSSFCell artCell = row.createCell(columnNum);
            HSSFCell nameCell = row.createCell(nameNum);


            if (am == null) {

                if (o == null) {
                    firstRowNum--;
                    continue;
                }

                if (o.toString().equals("")) {
                    artCell.setCellValue(l.get(0).toString());
                    nameCell.setCellValue(l.get(1).toString());
                    continue;
                }

                if ((Integer) o == 0) {
                    artCell.setCellValue(l.get(0).toString());
                    nameCell.setCellValue(l.get(1).toString());
                    row.createCell(firstAmNum).setCellValue(0);
                    continue;
                }
            }

            artCell.setCellValue(l.get(0).toString());
            nameCell.setCellValue(l.get(1).toString());
            for (int j = 0; j < am.size(); j++) {

                HSSFCell amCell = row.createCell(firstAmNum + j);
                amCell.setCellValue((Integer) am.get(j));

            }
        }

    }










    // - Overload,
    //  Write ArrayList of ArrayLists (every each contains String(0),Integer(1))
    // in order of our blank - first art, second amount
    // in particular existing book.
    // - Create sheet, and rows, if there is no!!
    // - Skip sell if there in no value for it
    public static void writeVertListInBookSingle(HSSFWorkbook book, String sheetName, int firstRowNum, int artColumnNum, int amColumnNum, ArrayList<ArrayList> orderList) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            System.out.println("sheet is null! Creating");
//            book.removeSheetAt(book.getSheetIndex(sheetName));
            sheet = book.createSheet(sheetName);
        }

        int firstLastRow = sheet.getLastRowNum();

        for (int i = 0; i < orderList.size(); i++) {

            ArrayList list = orderList.get(i);
            HSSFRow row;

            row = sheet.getRow(i + firstRowNum);
            if (row == null) {
                row = sheet.createRow(i + firstRowNum);
            }

            HSSFCell artCell = row.createCell(artColumnNum);
            HSSFCell amCell = row.createCell(amColumnNum);


            String art = (String) list.get(0);
            artCell.setCellValue(art);

            HashMap<String, String> nameMap = createMapOfArtsAndNamesPassport();
            if (amColumnNum - artColumnNum > 1) {
                //System.out.println("yes it is!");
                HSSFCell nameCell = row.createCell(artColumnNum + 1);
                nameCell.setCellValue(getNamePos(art, nameMap));
            }

            if (list.get(1) != null) {
                amCell.setCellValue((Integer) list.get(1));
                continue;
            } else {
                //amCell.setCellValue(0);
                continue;
            }


        }

        if (firstLastRow == sheet.getLastRowNum()) {

            for (int c = firstLastRow; c >= orderList.size(); c--) {

                if (sheet.getRow(c) == null) {
                    break;
                }
                sheet.removeRow(sheet.getRow(c));

            }

        }

        //System.out.println("Array is wroted!");

    }

    public static void writeVertListInBookSingleForReport(HSSFWorkbook book, String sheetName, ArrayList<ArrayList> orderList, Shipment ship) {

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("sheet is null! Creating");
            sheet = book.createSheet(sheetName);
            HSSFRow row1 = sheet.createRow(0);
            HSSFRow row2 = sheet.createRow(1);
            HSSFRow row3 = sheet.createRow(2);
            row3.createCell(2).setCellValue("Позиция");


            ArrayList artsAndNames = createListOfArtsAndNamesPassport();
            for (int i = 0; i < artsAndNames.size(); i++) {
                HSSFRow row = sheet.createRow(i + 3);
                ArrayList pos = (ArrayList) artsAndNames.get(i);
                row.createCell(1).setCellValue((String) pos.get(0));
                row.createCell(2).setCellValue((String) pos.get(1));
            }

            HSSFRow summRow = sheet.createRow(sheet.getLastRowNum() + 1);
        }


        int lastRowNum = sheet.getLastRowNum();

        SimpleDateFormat format = new SimpleDateFormat("dd MMMM");
        String data = format.format(ship.getShipDate());


        if (ship.getShipDate() == null) {

            System.out.println("Date is null in " + ship.getCustomer());
        }


        int colNum = sheet.getRow(2).getLastCellNum();
        sheet.getRow(1).createCell(colNum).setCellValue(ship.getCustomer());
        sheet.getRow(2).createCell(colNum).setCellValue(data);
        sheet.getRow(lastRowNum).createCell(colNum).setCellValue(ship.getSumm());

        ArrayList<Integer> newOrder = new ArrayList<Integer>();
        for (ArrayList pos : orderList) {
            int am;

            if (pos.get(1) == null) {
                am = 0;
            } else {
                am = (Integer) pos.get(1);
            }

            newOrder.add(am);

        }

        for (int i = 0; i < newOrder.size(); i++) {
            if (newOrder.get(i) == 0) {
                continue;
            }
            HSSFRow row = sheet.getRow(i + 3);
            row.createCell(colNum).setCellValue((Integer) newOrder.get(i));

        }


    }

    //Overload, write ArrayList of ArrayLists (String, Arraylist(Integer))(in order of our blank - first art, second amount) in particular existing book. Create sheet!. Create rows!
    public static void writeVertListInBookMultiple(HSSFWorkbook book, String sheetName, int firstRowNum, int artColumnNum, int amFirstColumnNum, ArrayList<ArrayList> orderList) {

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            sheet = book.createSheet(sheetName);

        }

        for (int i = 0; i < orderList.size(); i++) {

            HSSFRow row = sheet.createRow(i + firstRowNum);

            ArrayList list = orderList.get(i);

            HSSFCell artCell = row.createCell(artColumnNum);

            //String art = findArt(list.get(0).toString());
            String art = list.get(0).toString();
            artCell.setCellValue(art);

            HashMap<String, String> nameMap = createMapOfArtsAndNamesPassport();

            if (amFirstColumnNum - artColumnNum > 1) {
                HSSFCell nameCell = row.createCell(artColumnNum + 1);
                nameCell.setCellValue(getNamePos(art, nameMap));
            }


            ArrayList amList = (ArrayList) list.get(1);

            if (amList == null) {
                continue;
            }

            for (int r = 0; r < amList.size(); r++) {

                HSSFCell amCell = row.createCell(r + amFirstColumnNum);

                /*
                if (amCell == null) {
                    amCell = row.createCell(r + amFirstColumnNum);
                }
                */

                if ((Integer) amList.get(r) == 0) {
                    continue;
                }
                amCell.setCellValue((Integer) amList.get(r));

            }


        }

    }

    public static void writeHorListInBook(HSSFWorkbook book, String sheetName, int rowNum, int firstColumnNum, ArrayList list) {

        HSSFSheet sheet = book.getSheet(sheetName);

        HSSFRow row = sheet.getRow(rowNum);

        if (row == null) {row = sheet.createRow(rowNum);}

        int count = 0;

        for (int i = firstColumnNum; i < list.size() + firstColumnNum; i++) {

            HSSFCell cell = row.createCell(i);

            if (list.get(count) instanceof String) {
                // System.out.println("String, i = " + i + ", val " + list.get(i));
                cell.setCellValue((String) list.get(count));
            }

            if (list.get(count) instanceof Long) {
                // System.out.println("Long, i = " + i + ", val " + prices.get(i));
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue((Long) list.get(count));
            }

            if (list.get(count) instanceof Integer) {
                //  System.out.println("Integer, i = " + i + ", val " + prices.get(i));
                cell.setCellValue((Integer) list.get(count));
            }

            if (list.get(count) instanceof Double) {
                //  System.out.println("Double, i = " + i + ", val " + prices.get(i));
                cell.setCellValue((Double) list.get(count));
            }

            if (list.get(count) == null) {
                // System.out.println("Double, i = " + i + ", val " + list.get(i));
                cell.setCellValue("");
            }

            count++;
        }

    }

    public static void writeRemainsListInBookSingle(HSSFWorkbook book, String sheetName, ArrayList<ArrayList> list) {

        HSSFSheet sheet = book.getSheet(sheetName);
        HSSFRow rowFirst = sheet.createRow(sheet.getLastRowNum() + 2);
        HSSFCell cellFirst = rowFirst.createCell(0);

        cellFirst.setCellValue("ЗАКАЗАНО, НО ВЫВЕДЕНО");
        int header = cellFirst.getRowIndex() + 1;

        writeVertListInBookSingle(book, sheetName, header, 0, 2, list);

    }

    public static void writeRemainsListInBookMultiple(HSSFWorkbook book, String sheetName, ArrayList<ArrayList> list) {

        HSSFSheet sheet = book.getSheet(sheetName);
        if (sheet == null) {
            System.out.println("sheet is null! " + sheetName);
        }
        HSSFRow rowFirst = sheet.createRow(sheet.getLastRowNum() + 2);
        HSSFCell cellFirst = rowFirst.createCell(0);

        cellFirst.setCellValue("ЗАКАЗАНО, НО ВЫВЕДЕНО");
        int header = cellFirst.getRowIndex() + 1;

        writeVertListInBookMultiple(book, sheetName, header, 0, 2, list);

    }


    /*
    public static void writeRemainsListInBookMultiple(HSSFWorkbook book, String sheetName, ArrayList<ArrayList> list) {
        HSSFWorkbook outBook = createBook(getBlankFile());
        writeVertListInBookSingle(outBook, "Цены", 0, 3, price);
        HSSFSheet sheet = outBook.getSheet(outSheetName);

        for (int i = 9; i < outBook.getSheet(outSheetName).getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            String art = row.getCell(0).getStringCellValue();
            if (orderMap.containsKey(art)) {
                row.getCell(4).setCellValue(orderMap.get(art));
                orderMap.remove(art);
            } else {
                continue;
            }
        }
    }

    */


}


