package ToolsFiles.library;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.util.*;

import static ToolsFiles.library.StoreUtil.*;

public class StoreCreating {

    /////////////////// CREATE LIST FROM MAP


    // IN HashMap<Art, Amount>
    // IN ArrayList (arts)
    // OUT ArrayList <ArrayLists (art, am)
    // !!!!!!!!!!!SHORT
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL remove it from given map
    // ADDITIONAL print remains in console
    public static ArrayList<ArrayList> createShortListOrderFromSingleMap(HashMap<String, Integer> map, ArrayList<String> list) {

        ArrayList<ArrayList> newList = new ArrayList<ArrayList>();

        for (String art : list) {

            if (map.containsKey(art)) {
                ArrayList pos = new ArrayList();
                pos.add(art);
                pos.add(map.get(art));
                newList.add(pos);
                map.remove(art);
            }
        }

        if (map.size() > 0) {
            System.out.println("In remains map still is: ");
            printHashMap(map);

        }
        return newList;
    }

    // IN HashMap<Art, Amount>
    // IN ArrayList (arts)
    // OUT ArrayList <ArrayLists (art, am)
    // !!!!!!!!!!! LONG
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL remove it from given map
    // ADDITIONAL print remains in console
    // ADDITIONAL put 0
    public static ArrayList<ArrayList> createLongListOrderFromSingleMap(HashMap<String, Integer> map, ArrayList<String> list) {

        ArrayList<ArrayList> newList = new ArrayList<ArrayList>();

        for (String art : list) {

            int am = 0;

            if (map.containsKey(art)) {
                am = map.get(art);
                map.remove(art);
            }

            ArrayList pos = new ArrayList();
            pos.add(art);
            pos.add(am);
            //System.out.println("added " + art + ", " + am);
            newList.add(pos);
        }

        if (map.size() > 0) {
            System.out.println("In remains map still is: ");
            printHashMap(map);

        }

        return newList;
    }

    // IN HashMap<Art, T (arraylist or integer?)>
    // IN ArrayList (arts)
    // OUT ArrayList <ArrayLists (art, name, T)
    // !!!!!!!!!!!!SHORT
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL put NULL if amount is null
    // ADDITIONAL remove it from given map
    public static <T> ArrayList<ArrayList> createShortListOrderFromSingleMapWithNames(HashMap<String, T> map, ArrayList<ArrayList> list) {

        ArrayList<ArrayList> newList = new ArrayList<ArrayList>();

        for (ArrayList l : list) {

            String art = l.get(0).toString();

            if (map.containsKey(art)) {
                ArrayList pos = new ArrayList(3);
                String name = l.get(1).toString();
                T am = map.get(art);
                pos.add(art);
                pos.add(name);
                pos.add(am);
                newList.add(pos);

                map.remove(art);
            }

        }

        return newList;
    }

    // IN HashMap<Art, T (arraylist or integer?)>
    // IN ArrayList (arts)
    // OUT ArrayList <ArrayLists (art, name, T)
    // !!!!!!!!!!!!LONG LIST
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL put NULL if amount is null
    // ADDITIONAL remove it from given map
    public static <T> ArrayList<ArrayList> createLongListOrderFromSingleMapWithNames(HashMap<String, T> map, ArrayList<ArrayList> list) {

        ArrayList<ArrayList> newList = new ArrayList<ArrayList>();

        for (ArrayList l : list) {


            ArrayList pos = new ArrayList(3);
            String art = l.get(0).toString();
            String name = l.get(1).toString();
            T am = null;

            if (map.containsKey(art)) {
                am = map.get(art);
                map.remove(art);
            }

            pos.add(art);
            pos.add(name);
            pos.add(am);
            newList.add(pos);
        }

        return newList;
    }

    // IN HashMap<Art, T (array list or integer?)>
    // IN ArrayList (arts)
    // OUT ArrayList <ArrayLists (art, am)
    // !!!!!!!!!!!!SHORT
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL skip if amount is null
    // ADDITIONAL remove it from given map
    public static <T> ArrayList<ArrayList> createShortListOrderFromMultipleMap(HashMap<String, T> map, ArrayList<String> artList) {

        ArrayList<ArrayList> orderList = new ArrayList<ArrayList>();

        for (String art : artList) {

            if (map.containsKey(art)) {
                ArrayList ord = new ArrayList(2);
                T am = (T) map.get(art);
                ord.add(art);
                ord.add(am);
                orderList.add(ord);
                map.remove(art);
            }
        }
        return orderList;
    }

    // IN HashMap<Art, T (array list or integer?)>
    // IN ArrayList(arts)
    // OUT ArrayList <ArrayLists (art, Array)
    // !!!!!!!!!!!!LONG
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL put NULL if amount is null
    // ADDITIONAL remove it from given map
    public static <T> ArrayList<ArrayList> createLongListOrderFromMultipleMap(HashMap<String, T> map, ArrayList<String> artList) {

        ArrayList<ArrayList> orderList = new ArrayList<ArrayList>();

        for (String art : artList) {

            ArrayList ord = new ArrayList(2);

            T am = null;

            //System.out.println("map.get(art) " + map.get(art));

            if (map.containsKey(art)) {
                am = (T) map.get(art);
                //System.out.println("am " + am);
                map.remove(art);
            }

            ord.add(art);
            ord.add(am);
            //printArray(ord);
            orderList.add(ord);
        }
        return orderList;

    }

    // IN HashMap<Art, T (array list or integer?)>
    // IN ArrayList(arts)
    // OUT ArrayList <ArrayLists (art, , name, am)
    // !!!!!!!!!!!! SHORT
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL skip if amount is null
    // ADDITIONAL remove it from given map
    public static <T> ArrayList<ArrayList> createShortListOrderFromMultipleMapWithNames(HashMap<String, T> map, ArrayList<ArrayList> list) {

        ArrayList<ArrayList> newList = new ArrayList<ArrayList>();

        for (ArrayList l : list) {

            String art = l.get(0).toString();

            if (map.containsKey(art)) {
                String name = l.get(1).toString();
                T am = map.get(art);

                ArrayList pos = new ArrayList(3);
                pos.add(art);
                pos.add(name);
                pos.add(am);
                newList.add(pos);
                map.remove(art);
            }
        }
        return newList;
    }

    // IN HashMap<Art, T (array list or integer?)>
    // IN ArrayList(arts)
    // OUT ArrayList <ArrayLists (art, , name, ArrayList(am))
    // !!!!!!!!!!!!LONG
    // IN ORDER - blank or passport (up to list sended as argument)
    // ADDITIONAL put NULL if amount is null
    // ADDITIONAL remove it from given map
    public static <T> ArrayList<ArrayList> createLongListOrderFromMultipleMapWithNames(HashMap<String, T> map, ArrayList<ArrayList> list) {

        ArrayList<ArrayList> newList = new ArrayList<ArrayList>();

        for (ArrayList l : list) {

            String art = l.get(0).toString();
            String name = l.get(1).toString();
            T am = null;

            if (map.containsKey(art)) {
                am = map.get(art);
                map.remove(art);
            }

            ArrayList pos = new ArrayList(3);
            pos.add(art);
            pos.add(name);
            pos.add(am);
            newList.add(pos);

        }
        return newList;
    }


    /////////////////// CREATE ORDER

    // IN SINGLE order
    // IN ArrayList(art, name) from passport or blank
    // OUT ArrayList<Arraylist{art, name, am}>
    // !!!!!!!!!!!!!!!!!LONG
    // IN ORDER Blank or Passport - send it in arguments
    // ADDITIONAL:  if amounts is empty - set 0
    // ADDITIONAL:  remove positions from map
    public static ArrayList<ArrayList> createListOrderLongWithNames(HSSFWorkbook book, String sheetName, int firstRowNum, int artColNum, int amColNum, ArrayList<ArrayList> templateArray) {

        HashMap<String, Integer> map = createMapOrderSingle(book, sheetName, firstRowNum, artColNum, amColNum);

        ArrayList<ArrayList> orderList = new ArrayList<ArrayList>();

        for (ArrayList ar : templateArray) {

            ArrayList newAr = new ArrayList(3);

            newAr.add((String) ar.get(0));
            newAr.add((String) ar.get(1));

            int am = 0;

            if (map.containsKey(ar.get(0))) {
                am = map.get(ar.get(0));
                map.remove(ar.get(0));
            }

            newAr.add(am);

            orderList.add(newAr);
        }


        System.out.println("Не вошло в заказ");
        printHashMapWhithNames(map, createMapOfArtsAndNamesPassport());

        return orderList;
    }

    // IN SINGLE order
    // OUT ArrayList<Arraylist{art, am}>
    // !!!!!!!!!!!!!!!!! SHORT
    // IN ORDER like in order
    // ADDITIONAL:  if amounts is empty - skip it
    public static ArrayList<ArrayList> createListOrderShort(HSSFWorkbook book, String sheetName, int firstRowNum, int artNum, int amNum) {

        ArrayList<ArrayList> order = new ArrayList<ArrayList>();

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("Error! sheet " + sheetName + " is null!");
        }

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                System.out.println("Row " + i + " is null in sheet " + sheetName);
            }

            HSSFCell cellArt = row.getCell(artNum);

            if ((cellArt == null) || (cellArt.getStringCellValue() == "")) {
                continue;
            }

            String art = findArt(cellArt.getStringCellValue());

            int am = 0;
            HSSFCell cellAm = row.getCell(amNum);

            if (cellAm == null) {
                continue;
                //am = 0;
            } else {

                if (cellAm.getCellType() == CellType.NUMERIC) {
                    am = (int) cellAm.getNumericCellValue();
                }

                if ((cellAm.getCellType() == CellType.STRING) && (cellAm.getStringCellValue().length() > 0)) {
                    am = Integer.parseInt(cellAm.getStringCellValue());
                }

            }

            if (am == 0) {
                continue;
            }

            ArrayList pos = new ArrayList();
            pos.add(art);
            pos.add(am);
            order.add(pos);
        }

        return order;
    }

    // IN SINGLE order
    // OUT HASHMAP(art, am)>
    // !!!!!!!!!!!!!!!!! SHORT
    // IN ORDER no order
    // ADDITIONAL:  if amounts or art is empty - skip it
    public static HashMap<String, Integer> createMapOrderSingle(HSSFWorkbook book, String sheetName, int firstRowNum, int artColNum, int amColNum) {

        HashMap<String, Integer> map = new HashMap<String, Integer>();

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("ERROR! There is no sheet named '" + sheetName + "' in book!");
        }

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {continue;}

            if (row == null) {
                System.out.println("row is null!" + i);
            }

            HSSFCell artCell = row.getCell(artColNum);

            if (artCell == null) {
                //System.out.println("cell in row " + i + ", column " + artColNum + " is null, continue!");
                continue;
            }

            if (artCell.getStringCellValue() == null) {
                System.out.println("value in articul cell is null, continue!");
                continue;
            }

            if (artCell.getStringCellValue() == "") {
                //System.out.println("i = " + i + ", value is empty, continue!");
                continue;
            }

            String rowArt = artCell.getStringCellValue();

            String art = findArt(rowArt);


            if (art == "") {
                continue;
            }

            HSSFCell amCell = row.getCell(amColNum);

            Integer am;
            if ((amCell == null)||(amCell.getCellType() == CellType.STRING)) {
                //System.out.println("am is null! " + i);
                continue;
            } else {
                //amCell.setCellType(CellType.NUMERIC);
                am = (int) amCell.getNumericCellValue();
                if (am == 0) {
                    continue;
                }
            }

            if (map.containsKey(art)) {
                am = am + map.get(art);
            }

            map.put(art, am);
        }

        return map;
    }

    // Special version, dont do FindArt
    // OUT HASHMAP(art, am)>
    // !!!!!!!!!!!!!!!!! SHORT
    // IN ORDER no order
    // ADDITIONAL:  if amounts or art is empty - skip it
    public static HashMap<String, Integer> createMapOrderSingleForReport(HSSFWorkbook book, String sheetName, int firstRowNum, int artColNum, int amColNum) {

        HashMap<String, Integer> map = new HashMap<String, Integer>();

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("ERROR! There is no sheet named '" + sheetName + "' in book!");
        }

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                System.out.println("row is null!" + i);
            }

            HSSFCell artCell = row.getCell(artColNum);

            if (artCell == null) {
                System.out.println("cell in row " + i + ", column " + artColNum + " is null, continue!");
                continue;
            }

            if (artCell.getStringCellValue() == null) {
                System.out.println("value in articul cell is null, continue!");
                continue;
            }

            if (artCell.getStringCellValue() == "") {
                //System.out.println("i = " + i + ", value is empty, continue!");
                continue;
            }

            String rowArt = artCell.getStringCellValue();

            String art = rowArt;


            if (art == "") {
                continue;
            }

            HSSFCell amCell = row.getCell(amColNum);

            Integer am;
            if (amCell == null) {
                //System.out.println("am is null! " + i);
                continue;
            } else {
                amCell.setCellType(CellType.NUMERIC);
                am = (int) amCell.getNumericCellValue();
                if (am == 0) {
                    continue;
                }
            }

            map.put(art, am);
        }

        return map;
    }

    // IN MULTIPLE LINEAR order
    // OUT HASHMAP(art, ArrayList(am...)>
    // !!!!!!!!!!!!!!!!! SHORT
    // IN ORDER no order
    // ADDITIONAL:  if amounts is empty - add 0
    // ADDITIONAL:  have step for amounts values
    public static HashMap<String, ArrayList> createMapOrderMultiple(HSSFWorkbook book, String sheetName, int firstRowNum, int artColNum, int firstAmColNum, int step) {

        HashMap<String, ArrayList> map = new HashMap<String, ArrayList>();

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("Error : Sheet null " + sheetName);
            return null;
        }

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            HSSFCell artCell = row.getCell(artColNum);

            if (artCell == null) {
                System.out.println("i = " + i + ", Error: cell null, break!");
                continue;
            }

            String cellArt = artCell.getStringCellValue();

            if (cellArt == null) {
                System.out.println("i = " + i + ", Error: string value of cell is null, break!");
                break;
            }

            if (cellArt == "") {
                System.out.println("i = " + i + ", Error: string value of cell is empty, break!");
                break;
            }

            String art = findArt(cellArt);
            //System.out.println(art);
            if ((art == null)||(art.equals(""))) {
                System.out.println("Error: there is no articul in " + cellArt);
                continue;}

            ArrayList<Integer> amList = new ArrayList();

            for (int j = firstAmColNum; j <= row.getLastCellNum() - 1; j = j + step) {

                try {
                    HSSFCell amCell = row.getCell(j);
                    Integer am = (int) amCell.getNumericCellValue();
                    amList.add(am);

                } catch (Exception e) {
                    amList.add(0);
                    continue;
                }

                //System.out.println("Put " + art + ", " + am);
            }

            //System.out.println("Put " + art + ", " + amList);
            map.put(art, amList);
        }

        return map;
    }

    // IN MULTIPLE LINEAR order
    // OUT HASHMAP(art, ArrayList(am...)>
    // !!!!!!!!!!!!!!!!! SHORT
    // IN ORDER no order
    // ADDITIONAL:  if amounts is empty - add 0
    // ADDITIONAL:  have step for amounts values
    // ADDITIONAL:  DOES NOT FOUND ARTICUL IN KEY
    public static HashMap<String, ArrayList> createMapOrderMultipleSpecial(HSSFWorkbook book, String sheetName, int firstRowNum, int artColNum, int firstAmColNum, int step) {

        HashMap<String, ArrayList> map = new HashMap<String, ArrayList>();

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("Error : Sheet null " + sheetName);
            return null;
        }

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            HSSFCell artCell = row.getCell(artColNum);

            if (artCell == null) {
                System.out.println("i = " + i + ", Error: cell null, break!");
                continue;
            }

            String cellArt = artCell.getStringCellValue();

            if (cellArt == null) {
                System.out.println("i = " + i + ", Error: string value of cell is null, break!");
                break;
            }

            if (cellArt == "") {
                System.out.println("i = " + i + ", Error: string value of cell is empty, break!");
                break;
            }

            String art = cellArt;
            //System.out.println(art);
            if ((art == null)||(art.equals(""))) {
                System.out.println("Error: there is no articul in " + cellArt);
                continue;}

            ArrayList<Integer> amList = new ArrayList();

            for (int j = firstAmColNum; j <= row.getLastCellNum() - 1; j = j + step) {

                try {
                    HSSFCell amCell = row.getCell(j);
                    Integer am = (int) amCell.getNumericCellValue();
                    amList.add(am);

                } catch (Exception e) {
                    amList.add(0);
                    continue;
                }

                //System.out.println("Put " + art + ", " + am);
            }

            //System.out.println("Put " + art + ", " + amList);
            map.put(art, amList);
        }

        return map;
    }


    ///////////// COMMON

    //Create ArrayList of mutiple types, in case of blank cell add NULL
    public static ArrayList createVertListOfValues(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum) {

        ArrayList list = new ArrayList();
        HSSFSheet sheet = book.getSheet(sheetName);

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                System.out.println("i " + i + ", row is null, break!");
                break;
            }

            if (sheet.getRow(i).getCell(columnNum) == null) {
                //System.out.println("i " + i + ", cell is null, break!");
                break;
            }

            HSSFCell cell = row.getCell(columnNum);

            switch (cell.getCellType()) {

                case BLANK:
                    list.add(null);
                    break;
                case STRING:
                    String s = cell.getStringCellValue();
                    if (s.length() > 0) {
                        list.add(s);
                    }
                    break;
                case NUMERIC:
                    double q = cell.getNumericCellValue();

                    double r = q % (int) q;
                    if (r == 0) {
                        int w = (int) q;
                        list.add(w);
                    } else {
                        list.add(q);
                    }

                    break;

            }

        }

        return list;

    }

    //Create ArrayList of mutiple types, in case of blank cell add Object (0, null or "", as you send)
    public static ArrayList createVertListOfValues(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, Object o) {

        ArrayList list = new ArrayList();
        HSSFSheet sheet = book.getSheet(sheetName);
        //sheet.ungroupRow(0, sheet.getLastRowNum());

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            if (sheet.getRow(i) == null) {
                //System.out.println("i " + i + ", row is null, break!");
                break;
            }

            if (sheet.getRow(i).getCell(columnNum) == null) {
                //System.out.println("i " + i + ", cell is null, break!");
                break;
            }

            HSSFCell cell = sheet.getRow(i).getCell(columnNum);

            switch (cell.getCellType()) {

                case BLANK:
                    list.add(o);
                    break;
                case STRING:
                    String s = cell.getStringCellValue();
                    if (s.length() > 0) {
                        list.add(s);
                    }
                    break;
                case NUMERIC:
                    double q = cell.getNumericCellValue();

                    double r = q % (int) q;
                    if (r == 0) {
                        int w = (int) q;
                        list.add(w);
                    } else {
                        list.add(q);
                    }

                    break;

            }

        }

        return list;

    }

    //Create ArrayList of mutiple types, in case of blank cell skip it, if cell or row = null - skip it too
    public static ArrayList createVertListOfValues(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, boolean b) {

        ArrayList list = new ArrayList();
        HSSFSheet sheet = book.getSheet(sheetName);
        //sheet.ungroupRow(0, sheet.getLastRowNum());

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            if (sheet.getRow(i) == null) {
                System.out.println("i " + i + ", row is null, break!");
                continue;
            }

            if (sheet.getRow(i).getCell(columnNum) == null) {
                //System.out.println("i " + i + ", cell is null, break!");
                continue;
            }

            HSSFCell cell = sheet.getRow(i).getCell(columnNum);

            switch (cell.getCellType()) {

                case BLANK:
                    break;
                case STRING:
                    String s = cell.getStringCellValue();
                    if (s.length() > 0) {
                        list.add(s);
                    }
                    break;
                case NUMERIC:
                    double q = cell.getNumericCellValue();

                    double r = q % (int) q;
                    if (r == 0) {
                        int w = (int) q;
                        list.add(w);
                    } else {
                        list.add(q);
                    }

                    break;

            }

        }

        return list;

    }

    //Create ArrayList of DOUBLE types, in case of blank cell skip it, if cell or row = null - skip it too
    public static ArrayList createVertListOfValuesDouble(HSSFWorkbook book, String sheetName, int firstRowNum, int columnNum, boolean b) {

        ArrayList list = new ArrayList();
        HSSFSheet sheet = book.getSheet(sheetName);
        //sheet.ungroupRow(0, sheet.getLastRowNum());

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            if (sheet.getRow(i) == null) {
                System.out.println("i " + i + ", row is null, break!");
                continue;
            }

            if (sheet.getRow(i).getCell(columnNum) == null) {
                //System.out.println("i " + i + ", cell is null, break!");
                continue;
            }

            HSSFCell cell = sheet.getRow(i).getCell(columnNum);

            switch (cell.getCellType()) {

                case BLANK:
                    break;
                case STRING:
                    String s = cell.getStringCellValue();
                    if (s.length() > 0) {
                        list.add(s);
                    }
                    break;
                case NUMERIC:
                    double q = cell.getNumericCellValue();
                    list.add(q);
                    break;

            }

        }

        return list;

    }

    public static ArrayList removeNullsFromArray(ArrayList list) {

        for (Object o : list) {

            if (o == null) {
                list.remove(o);
            }

            if ((o instanceof Integer) && ((Integer) o == 0)) {
                list.remove(o);
            }

            if ((o instanceof String) && (o.equals(""))) {
                list.remove(o);
            }

            if ((o instanceof Double) && ((Double) o == 0.0)) {
                list.remove(o);
            }
        }

        return list;
    }

    //Out: ArrayList of values (String, double or int)
    public static ArrayList createHorListOfValues(HSSFWorkbook book, String sheetName, int rowNum, int firstColumnNum) {

        ArrayList list = new ArrayList();

        HSSFRow row = book.getSheet(sheetName).getRow(rowNum);

        if (row == null) {
            System.out.println("Row " + rowNum + " is null!");
        }

        for (int i = firstColumnNum; i < row.getLastCellNum(); i++) {

            HSSFCell cell = row.getCell(i);

            if (cell == null) {
                //System.out.println("Cell " + i + " is null, break!");
                break;
            }

            switch (cell.getCellType()) {

                case STRING:
                    String s = cell.getStringCellValue();
                    if (s.length() > 0) {
                        list.add(s);
                    }
                    break;

                case NUMERIC:
                    double q = cell.getNumericCellValue();

                    double r = q - (int) q;
                    if (r == 0) {
                        int w = (int) q;
                        list.add(w);
                    } else {
                        list.add(q);
                    }

                    break;

                    default:
                        //System.out.println(cell.getCellType());
                        String e = cell.getStringCellValue();
                        if (e.length() > 0) {
                            list.add(e);
                        }
                        break;

            }

        }

        return list;

    }


    /////////////////////   -----------------FROM BLANK
    //Create ArrayList<arts> from actual blank
    public static ArrayList<String> createListOfArtsFromBlank() {

        ArrayList<String> list;
        HSSFWorkbook blankBook = createBook(getBlankFile());
        list = createVertListOfValues(blankBook, "Форма", 9, 0, true);
        closeBook(blankBook);
        return list;
    }

    //Create ArrayList<ArrayList (art, name)> - from blank
    public static ArrayList<ArrayList> createListOfArtsAndNamesBlank() {

        ArrayList<ArrayList> goodsList = new ArrayList<ArrayList>();

        HSSFWorkbook book = createBook(getBlankFile());

        HSSFSheet sheet = book.getSheet("Форма");

        for (int i = 9; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            HSSFCell artCell = row.getCell(0);
            String art = artCell.getStringCellValue();
            if (art == "") {
                continue;
            }

            HSSFCell nameCell = row.getCell(1);
            String name = nameCell.getStringCellValue();
            //System.out.println("Put Key " + art + ", Value " + name);
            ArrayList pos = new ArrayList(2);
            pos.add(art);
            pos.add(name);
            goodsList.add(pos);
        }

        closeBook(book);

        return goodsList;
    }

    //Create Map String of Articuls - Names from Blank
    public static HashMap<String, String> createMapOfArtsAndNamesBlank() {

        HashMap<String, String> goodsMap = new HashMap<String, String>();

        HSSFWorkbook book = createBook(getBlankFile());

        HSSFSheet sheet = book.getSheet("Форма");

        if (sheet == null) {
            System.out.println("Sheet is null!");
        }

        for (int i = 9; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                System.out.println("Row is null!");
                continue;
            }

            HSSFCell artCell = row.getCell(0);
            String art = artCell.getStringCellValue();

            HSSFCell nameCell = row.getCell(1);
            String name = nameCell.getStringCellValue();
            //System.out.println("Put Key " + art + ", Value " + name);
            goodsMap.put(art, name);
        }

        closeBook(book);

        return goodsMap;
    }

    //Create Map String of Articuls - Names from Blank
    public static HashMap<String, String> createMapOfArtsAndNamesBlankSpecial() {

        HashMap<String, String> goodsMap = new HashMap<String, String>();

        HSSFWorkbook book = createBook(getWorkDir() + "ТФ шоколад артикулы.xls");

        HSSFSheet sheet = book.getSheet("shok art");

        if (sheet == null) {
            System.out.println("Sheet is null!");
        }

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                System.out.println("Row is null!");
                continue;
            }

            HSSFCell nameCell = row.getCell(0);
            String nameGoods = nameCell.getStringCellValue();

            HSSFCell artCell = row.getCell(1);
            String artGoods = artCell.getStringCellValue();
            goodsMap.put(nameGoods, artGoods);
        }

        closeBook(book);

        return goodsMap;
    }

    /////////////////////  ---------------FROM PASSPORT
    //Create ArrayList<arts> from from passport
    public static ArrayList<String> createListOfArtsFromPassport() {

        ArrayList<String> list;
        HSSFWorkbook blankBook = createBook(getPassportFile());
        list = createVertListOfValues(blankBook, "Форма", 0, 0);
        closeBook(blankBook);
        return list;
    }

    //Create ArrayList<ArrayList (art, name)> - from passport
    public static ArrayList<ArrayList> createListOfArtsAndNamesPassport() {

        ArrayList<ArrayList> goodsList = new ArrayList<ArrayList>();

        HSSFWorkbook book = createBook(getPassportFile());

        HSSFSheet sheet = book.getSheet("Форма");

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            HSSFCell artCell = row.getCell(0);
            String art = artCell.getStringCellValue();

            HSSFCell nameCell = row.getCell(1);
            String name = nameCell.getStringCellValue();
            //System.out.println("Put Key " + art + ", Value " + name);
            ArrayList pos = new ArrayList(2);
            pos.add(art);
            pos.add(name);
            goodsList.add(pos);
        }

        closeBook(book);

        return goodsList;
    }

    //Create Map String of Articuls - Names from passport
    public static HashMap<String, String> createMapOfArtsAndNamesPassport() {

        HashMap<String, String> goodsMap = new HashMap<String, String>();

        HSSFWorkbook book = createBook(getPassportFile());

        HSSFSheet sheet = book.getSheet("Форма");

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            HSSFCell artCell = row.getCell(0);
            String art = artCell.getStringCellValue();

            HSSFCell nameCell = row.getCell(1);
            String name = nameCell.getStringCellValue();
            //System.out.println("Put Key " + art + ", Value " + name);
            goodsMap.put(art, name);
        }

        closeBook(book);

        return goodsMap;
    }




    /////////////////////  ---------------FROM Clients and Prices

    //Create HashMap<Long, String> from from cl and pr
    public static HashMap<Long, String> createMapOfInnFromClAndPr() {
        HSSFWorkbook book = createBook(getClientsAndPricesFile());
        HSSFSheet sheet = book.getSheet("ИНН");
        int l = sheet.getLastRowNum();
        //System.out.println("Last row is " + l);
        HashMap<Long, String> innMap = new HashMap<Long, String>();

        for (int i = 0; i <= l; i++) {

            String name = sheet.getRow(i).getCell(0).getStringCellValue();
            Long inn = (long) sheet.getRow(i).getCell(1).getNumericCellValue();
            innMap.put(inn, name);
        }

        closeBook(book);
        return innMap;
    }

    public static HashSet<String> createSetOfCashNamesFromClAndPr() {
        HSSFWorkbook book = createBook(getClientsAndPricesFile());
        HSSFSheet sheet = book.getSheet("ДЖК");
        int l = sheet.getLastRowNum();
        HashSet<String> nameSet = new HashSet<String>();

        for (int i = 0; i <= l; i++) {

            String name = sheet.getRow(i).getCell(0).getStringCellValue();
            nameSet.add(name);
        }

        closeBook(book);
        return nameSet;
    }

    //Create Array String of client names from clAndPr from sheet DATA
    public static ArrayList<String> createListOfClientsFromData() {

        HSSFWorkbook clientsBook = createBook(getClientsAndPricesFile());
        ArrayList list = createVertListOfValues(clientsBook, "Данные", 1, 6);
        closeBook(clientsBook);

        ListIterator iter = list.listIterator();

        while (iter.hasNext()) {

            System.out.println("\nstart");

            System.out.println("Index: " + list.indexOf(iter.hasNext()));

            System.out.println(iter.next());
            //System.out.println(String.valueOf(iter.next()));
            System.out.println("Class: " + iter.next().getClass());

            if (iter.next().toString() == "0") {
                System.out.println("is 0!");
                //iter.remove();}
            }
            System.out.println("fin\n");
        }


        return list;
    }


}

