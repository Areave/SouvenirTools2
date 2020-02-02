package ToolsFiles;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.*;

import static ToolsFiles.library.StoreUtil.*;
import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreWriting.*;
import static ToolsFiles.library.Test.*;

public class Main {

    public static void main(String[] args) {

        //library.Test.main();

        //makeClientOrderBlank();


        //------------Create order---------------------

        //createSingleOrderFromOldOrderBlank(getSingleOrderFromOldOrderBlankDir());

        //createSingleOrderFromClientOrder(getSingleOrderFromClientOrderDir());




        //in - unsorted tf choc order
        //out - sorted (blank) multiple order in same file sheet "Форма"
        //mainForTfOrderChok(getMultipleOrderFromClientOrderDir(), "tf choc.xls");

        //createMultipleOrderFromClientOrder(getMultipleOrderFromClientOrderDir());


        //--------------------Create report------------------------

        makeFullCycleOfShipments ("2020","1");

        //comparasingShipment("2018", "2019", 10, 3);


        //in directory, given as a parameter
        // need to be files mag.xls and chok.xls
        //createTFFullReport(getTfReportsDir());

        //Put Бука and Пароль НН (edited! 25_9) and ТФ Сувенир (prepared! by method above)
        // reports for exactly needed month
        // and others from Алена
        // in CommonSalesReport folder in acceptable format
        //makeCommonSalesReport(getWorkDir() + getCommonSalesReportDir());


        //--------------------Util----------------------

        // Use "temp" folder
        // Parse prices from "Себестоимость" file
        // put price and summ in every act in folder
        // and print scipped articuls
        //putPricesToReturnActs();

        //createSortedOrder(getSortedOrderDir());

        //updateFileForNewPassport(getUpdateFileForNewPassportDir());

        //createSalesPlanForDimon();

        // Use file "reestr" from workDir
        // need a number of kvartal
        // writing 1
        //reestrSimple(4);

        //write remains from actual remains file "remains.xls" in main work dirictory
        // in blank in "Remains" directory
        //putRemainsInOrderBlank(getRemainsDir());

        //take a number of month,
        // print every sheet name,
        // that have no digits in summ sales cell
        // in sale report
        //library.Test.findClientsWithEmptyReports(12);


        //---------------------Deprecated---------------------

        //Work with files from same directory
        // - DEPRECATED!
        //makeCommonSalesReportAlt(getWorkDir() + getCommonSalesReportDir());


        //library.Test.mainForTfReportChok();

    }

    public static void makeFullCycleOfShipments(String year, String month) {

        String workDir = getWorkDir() + "Shipment/" + year + "/" + month + "/";

        checkDirOfShipmentsForErrors(workDir);
        //createAllFilesDirectoryForMonth(workDir);
        //putInnFromEveryShipsInDirToClAndPrices(workDir + "allFiles/");
        //tempPutNameOfClientFromEveryShipsInDirToClAndPrices(year,Integer.parseInt(month));
        //library.ShipmentStore.main("Shipment/" + year + "/" + month + "/cashFiles/", "Отчет отгрузки 2020.xls", 2);
        //library.ShipmentStore.main("Shipment/" + year + "/" + month + "/bankFiles/", "Отчет отгрузки 2020.xls", 2);
        //createAllFilesDirectoryForMonth(workDir);
        //createReportSummFile(year, Integer.parseInt(month), Integer.parseInt(month));

    }

    public static void createTFFullReport(String workDirectory) {

        //create report file in report directory
        String tfReportFileName = getWorkDir() + getCommonSalesReportDir() + "ТФ Сувенир.xls";
        File tfReportFile = new File(tfReportFileName);

        if (tfReportFile.exists()) {
            System.out.println("TF Report file exist, delete");
            tfReportFile.delete();
        }

        HSSFWorkbook tfReportBook = new HSSFWorkbook();
        HSSFSheet sheet = tfReportBook.createSheet("report");

        try (FileOutputStream out = new FileOutputStream(tfReportFile)) {
            tfReportBook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("TF Report File has been created");

        // put data from mag.xls to report file

        HSSFWorkbook magBook = createBook(getWorkDir() + workDirectory + "mag.xls");

        writeBook(magBook, getWorkDir() + workDirectory + "mag.xls");

        //then put data
        HashMap<String, Integer> map = createMapOrderSingle(magBook, "Sheet1", 7, 0, 5);
        closeBook(magBook);

        ArrayList<String> maskArray = createListOfArtsFromPassport();

        ArrayList orderList = createLongListOrderFromMultipleMap(map, maskArray);

        // Write it to out book

        writeVertListInBookSingle(tfReportBook, "report", 0, 0, 2, orderList);

        //Create List of remains in order of Passport
        //ArrayList<ArrayList> remainsArray = createShortListOrderFromSingleMap(map, createListOfArtsFromPassport());
        //writeRemainsListInBookMultiple(outbook, orderName, remainsArray);
        if (!map.isEmpty()) {
            System.out.println("Map is not empty!");
            printHashMap(map);
        }

        writeBook(tfReportBook, tfReportFileName);
        closeBook(tfReportBook);


        //---------------------------------------------------
        // put data from chok.xls to report file
        //

        HSSFWorkbook chokBook = createBook(getWorkDir() + workDirectory + "chok.xls");


        //prepare chok file

        HashMap<String, ArrayList> orderMap = createMapOrderMultipleSpecial(chokBook, "Sheet1", 6, 0, 2, 1);

        closeBook(chokBook);

        HashMap<String, String> namesMap = createMapOfArtsAndNamesBlankSpecial();

        HashMap<String, Integer> newMap = new HashMap<String, Integer>();

        for (Map.Entry<String, ArrayList> entry : orderMap.entrySet()) {

            String goodsName = entry.getKey();

            if (namesMap.containsKey(goodsName)) {
                ArrayList<Integer> list = entry.getValue();
                Integer val = list.get(0);
                newMap.put(namesMap.get(goodsName), val);
            }

        }

        // now we have a ordinary newMap<String, Integer>
        // and need to write it to report file
        // but not to delete previus data

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            String art = sheet.getRow(i).getCell(0).getStringCellValue();
            if (newMap.containsKey(art)) {
                sheet.getRow(i).createCell(2).setCellValue(newMap.get(art));
                newMap.remove(art);
            }
        }

        if (!newMap.isEmpty()) {
            System.out.println("Map is not empty!");
            printHashMap(newMap);
        }

        writeBook(tfReportBook, getWorkDir() + getCommonSalesReportDir() + "ТФ Сувенир.xls");
        closeBook(tfReportBook);

        openFile(getWorkDir() + getCommonSalesReportDir() + "ТФ Сувенир.xls");

    }

    // IN: file 'remains.xls' from main work directory
    // IN: blank file in 'Remain' directory
    // OUT: take numbers of remains, put it to given blank
    public static void putRemainsInOrderBlank(String thisWorkDirectory) {

        //Create map of remains
        HSSFWorkbook book = createBook(getWorkDir() + "remains.xls");
        HashMap<String, Integer> remainsMap = createMapOrderSingle(book, "TDSheet", 12, 2, 3);
        closeBook(book);

        //walk through blank, writing numbers

        ArrayList<String> list = createDirectoryList(getWorkDir() + thisWorkDirectory);

        System.out.println("Выберите файл, в который нужно внести остатки");

        int n = chooseElementFromList(list);

        String fileName = list.get(n);

        String fileFullName = getWorkDir() + thisWorkDirectory + "/" + fileName;

        HSSFWorkbook book2 = createBook(fileFullName);

        for (int i = 9; ; i++) {

            HSSFRow row = book2.getSheet("ФОРМА").getRow(i);

            if (row == null) {
                break;
            }

            HSSFCell cell = row.getCell(0);

            if (cell == null) {
                break;
            }

            String art = cell.getStringCellValue();

            if (remainsMap.containsKey(art)) {

                book2.getSheet("ФОРМА").getRow(i).createCell(14).setCellValue(remainsMap.get(art));
                remainsMap.remove(art);
            }

        }

        printHashMap(remainsMap);

        writeBook(book2, fileFullName);
        closeBook(book2);
        openFile(fileFullName);

    }

    //need save file in work dir with name "reestr.xls"
    // rename sheet as "все"
    // define last row in code below
    // e - exit
    public static void reestr() {

        HSSFWorkbook book = createBook(getWorkDir() + "reestr.xls");


        //create map of docs
        HashMap<Integer, String> map = new HashMap();
        HashMap<Integer, String> mapShort = new HashMap();
        ArrayList<Map.Entry<Integer, String>> list = new ArrayList<Map.Entry<Integer, String>>(3);

        int lastRow = 1431;


        ArrayList<HashMap> mapList = new ArrayList<HashMap>();
        HashMap<Integer, String> map1 = new HashMap();
        HashMap<Integer, String> map2 = new HashMap();
        HashMap<Integer, String> map3 = new HashMap();
        HashMap<Integer, String> map4 = new HashMap();
        mapList.add(map1);
        mapList.add(map2);
        mapList.add(map3);
        mapList.add(map4);

        for (int k = 1; k < 5; k++) {
            HashMap hm = mapList.get(k - 1);
            HSSFSheet sheet = book.getSheet(k + " квартал");

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {

                HSSFRow row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                HSSFCell cell = row.getCell(1);
                if (cell == null) {
                    continue;
                }

                String d = row.getCell(1).getStringCellValue();
                hm.put(i, d);
            }

        }

        closeBook(book);

        //endless cycle

        for (; ; ) {
            //get a TN number
            String num, com;
            int n = -1;
            Scanner s = new Scanner(System.in);

            System.out.println("Введите номер накладной");
            num = s.next();

            if (num.equals("e") || (num.equals("е"))) {
                break;
            }

            if (num.length() == 3) {
                num = "0" + num;
            } else if (num.length() == 2) {
                num = "00" + num;
            } else if (num.length() == 1) {
                num = "000" + num;
            }

            int numInt = Integer.parseInt(num);

            //check if there is a similar number in map and find a number of row

            for (HashMap<Integer, String> partMap : mapList) {

                for (Map.Entry<Integer, String> entry : partMap.entrySet()) {

                    String value = entry.getValue();
                    String valueShort = value.substring(7, 11);

                    if (valueShort.equals(num)) {

                        int key = entry.getKey();

                        if (mapShort.containsKey(key)) {

                            key = key + 1000000;
                        }

                        mapShort.put(key, value);
                    }
                }

            }

            if (mapShort.isEmpty()) {
                System.out.println("Нет такого номера!");
                continue;
            }

            if (mapShort.size() == 1) {

                for (Map.Entry<Integer, String> e : mapShort.entrySet()) {

                    n = e.getKey();
                    if (n > 100000) {
                        n = n - 1000000;
                    }
                    mapShort.clear();
                }
            }

            if (mapShort.size() > 1) {

                System.out.println("Выберите номер верной отгрузки");

                list.add(null);
                list.add(null);
                list.add(null);

                for (Map.Entry<Integer, String> e : mapShort.entrySet()) {

                    String v = e.getValue();
                    String c = v.substring(0, 1);

                    if (c.equals("И")) {
                        list.set(0, e);
                    } else if (c.equals("П")) {
                        list.set(1, e);
                    } else {
                        list.set(2, e);
                    }

                }

                for (int r = 0; r < 3; r++) {

                    Map.Entry<Integer, String> m = list.get(r);

                    if (m == null) {
                        continue;
                    }
                    System.out.println(r + ", " + m.getValue());

                }

                n = s.nextInt();

                Map.Entry<Integer, String> m = list.get(n);

                n = m.getKey();
                if (n > 100000) {
                    n = n - 1000000;
                }
                mapShort.clear();

                list.clear();


            }

            com = map.get(n).substring(0, 2);
            HSSFWorkbook book2 = createBook(getWorkDir() + "reestr.xls");
            HSSFSheet sheet2 = book2.getSheet("все");
            String y = sheet2.getRow(n).getCell(3).getStringCellValue();
            sheet2.getRow(n).createCell(7).setCellValue("1");
            System.out.println("Записано: " + com + ", " + y + ", номер строки " + String.valueOf(n + 1));
            map.remove(n);
            writeBook(book2, getWorkDir() + "reestr.xls");
            closeBook(book2);

        }

    }

    public static void reestrSimple(int kvartalNum) {

        HSSFWorkbook book = createBook(getWorkDir() + "reestr.xls");

        //create map of docs
        HashMap<Integer, String> map = new HashMap();
        HashMap<Integer, String> mapShort = new HashMap();
        ArrayList<Map.Entry<Integer, String>> list = new ArrayList<Map.Entry<Integer, String>>(4);

        //ArrayList<HashMap> mapList = new ArrayList<HashMap>();
        HashMap<Integer, String> mapLongTTNNum = new HashMap();
        //mapList.add(mapLongTTNNum);

        HSSFSheet sheet = book.getSheet(kvartalNum + " квартал");

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            HSSFCell cell = row.getCell(1);
            if (cell == null) {
                continue;
            }

            String d = row.getCell(1).getStringCellValue();
            mapLongTTNNum.put(i, d);
        }

        closeBook(book);

        System.out.println("Внесение единиц в файл с реестром! квартал " + kvartalNum);

        //endless cycle

        for (; ; ) {
            //get a TN number
            String num, com;
            int n = -1;
            Scanner s = new Scanner(System.in);
            System.out.println();
            System.out.println("Введите номер накладной");
            num = s.next();

            if (num.equals("e") || (num.equals("е"))) {
                break;
            }

            if (num.length() == 3) {
                num = "0" + num;
            } else if (num.length() == 2) {
                num = "00" + num;
            } else if (num.length() == 1) {
                num = "000" + num;
            }

            //check if there is a similar number in map and find a number of row

            for (Map.Entry<Integer, String> entry : mapLongTTNNum.entrySet()) {

                String ttnNum = entry.getValue();
                String ttnNumShort = ttnNum.substring(7, 11);

                if (ttnNumShort.equals(num)) {

                    int rowNum = entry.getKey();

                    //System.out.println("rowNum " + rowNum);

                    if (mapShort.containsKey(rowNum)) {

                        rowNum = rowNum + 1000000;

                    }

                    mapShort.put(rowNum, ttnNum);
                }
            }

            if (mapShort.isEmpty()) {
                System.out.println("Нет такого номера!");
                continue;
            }

            if (mapShort.size() == 1) {

                for (Map.Entry<Integer, String> e : mapShort.entrySet()) {

                    n = e.getKey();
                    if (n > 100000) {
                        n = n - 1000000;
                    }
                    mapShort.clear();
                }
            }

            if (mapShort.size() > 1) {

                System.out.println("Выберите номер верной отгрузки");

                list.add(null);
                list.add(null);
                list.add(null);
                list.add(null);

                for (Map.Entry<Integer, String> e : mapShort.entrySet()) {

                    String v = e.getValue();
                    String c = v.substring(0, 1);

                    if (c.equals("И")) {
                        list.set(0, e);
                    } else if (c.equals("П")) {
                        list.set(1, e);
                    } else if (c.equals("Д")) {
                        list.set(2, e);
                    } else {
                        list.set(3, e);
                    }

                }

                for (int r = 0; r <= 3; r++) {

                    Map.Entry<Integer, String> m = list.get(r);

                    if (m == null) {
                        continue;
                    }
                    System.out.println(r + ", " + m.getValue());

                }

                n = s.nextInt();

                Map.Entry<Integer, String> m = list.get(n);

                n = m.getKey();
                if (n > 100000) {
                    n = n - 1000000;
                }
                mapShort.clear();

                list.clear();


            }

            com = mapLongTTNNum.get(n).substring(0, 2);

            HSSFWorkbook book2 = createBook(getWorkDir() + "reestr.xls");
            HSSFSheet sheet2 = book2.getSheet(kvartalNum + " квартал");
            String y = sheet2.getRow(n).getCell(3).getStringCellValue();
            sheet2.getRow(n).createCell(7).setCellValue(1);
            System.out.println("Записано: " + com + ", " + y + ", номер строки " + String.valueOf(n + 1));
            mapLongTTNNum.remove(n);
            writeBook(book2, getWorkDir() + "reestr.xls");
            closeBook(book2);

        }

    }

    public static void makeCommonSalesReport(String thisWorkDir) {

        ArrayList<String> dirAr = createDirectoryList(thisWorkDir);

        dirAr.remove("Отчет продажи 2020.xls");
        dirAr.remove("Архив");
        printArray(dirAr);


        Scanner s = new Scanner(System.in);
        System.out.println("Введите номер считываемого столбца");
        //int n = s.nextInt();
        int n = 25;

        for (String filename : dirAr) {

            if (!(filename.equals("ТФ Сувенир.xls"))) {
                continue;
            }

            HSSFWorkbook book = createBook(thisWorkDir + filename);

            System.out.println("Обрабатывается файл: " + filename);

            HashMap<String, Integer> map;

            if (filename.equals("Бука.xls")) {
                map = createMapOrderSingle(book, "выгрузка накладных по выбранном", 1, 5, 2);
                closeBook(book);
            } else if (filename.equals("Пароль НН.xls")) {
                map = createMapOrderSingle(book, "сводка по наименованиям с групп", 3, 7, 10);
                closeBook(book);
            } else if (filename.equals("ТФ Сувенир.xls")) {
                map = createMapOrderSingle(book, "report", 0, 0, 2);
                closeBook(book);
            } else {
                map = createMapOrderSingle(book, "2019", 3, 0, n);
                closeBook(book);
            }


            HSSFWorkbook reportBook = createBook(thisWorkDir + "Отчет продажи 2020 .xls");

            String sheetName = filename.substring(0, filename.length() - 4);

            HSSFSheet sheet = reportBook.getSheet(sheetName);

            //System.out.println("Last row is " + sheet.getLastRowNum());

            for (int i = 2; i < sheet.getLastRowNum(); i++) {

                if (sheet.getRow(i) == null) {
                    System.out.println("Row " + i + " is null");
                }

                if (sheet.getRow(i).getCell(0) == null) {
                    System.out.println("Cell on row " + i + " is null");
                }

                String art = sheet.getRow(i).getCell(0).getStringCellValue();

                if (map.containsKey(art)) {

                    int am = map.get(art);
                    sheet.getRow(i).createCell(n).setCellValue(am);
                    map.remove(art);

                }

            }

            if (!map.isEmpty()) {

                System.out.println(sheetName + ", map is not empty! size is " + map.size());
                printHashMap(map);
            }

            writeBook(reportBook, thisWorkDir + "Отчет продажи 2020.xls");
            closeBook(reportBook);

        }

        openFile(thisWorkDir + "Отчет продажи 2020.xls");
    }

    public static void makeCommonSalesReportAlt(String thisWorkDir) {

        String thisWorkDirFiles = thisWorkDir + "Бука/";

        ArrayList<String> dirAr = createDirectoryList(thisWorkDirFiles);

        printArray(dirAr);

        Scanner s = new Scanner(System.in);
        System.out.println("Введите номер считываемого столбца");
        int m = s.nextInt();

        for (String filename : dirAr) {

            HSSFWorkbook book = createBook(thisWorkDirFiles + filename);

            System.out.println("Обрабатывается файл: " + filename);

            HashMap<String, Integer> map = createMapOrderSingle(book, "выгрузка накладных по выбранном", 1, 5, 2);
            closeBook(book);

            System.out.println("Сумма: " + mapValueSumm(map));


            HSSFWorkbook reportBook = createBook(thisWorkDir + "Отчет продажи.xls");

            String sheetNameBlank = "Общий отчет";
            String sheetNameReport = "Бука";

            HSSFSheet sheetBlank = reportBook.getSheet(sheetNameBlank);
            HSSFSheet sheetReport = reportBook.getSheet(sheetNameReport);

            //System.out.println("Last row is " + sheet.getLastRowNum());

            for (int i = 2; i < sheetBlank.getLastRowNum(); i++) {

                if (sheetBlank.getRow(i) == null) {
                    System.out.println("Row " + i + " is null");
                }

                if (sheetBlank.getRow(i).getCell(0) == null) {
                    System.out.println("Cell on row " + i + " is null");
                }

                String art = sheetBlank.getRow(i).getCell(0).getStringCellValue();

                if (map.containsKey(art)) {
                    int am = map.get(art);
                    sheetReport.getRow(i).createCell(m).setCellValue(am);
                    map.remove(art);

                }

            }

            System.out.println("Summ is now " + mapValueSumm(map));

            if (!map.isEmpty()) {

                System.out.println(sheetNameReport + ", map is not empty! size is " + map.size());
                printHashMap(map);
            }

            writeBook(reportBook, thisWorkDir + "Отчет продажи.xls");
            closeBook(reportBook);

        }

        openFile(thisWorkDir + "Отчет продажи.xls");
    }


    // Use console for printing list of clients
    // Make client blank, open it
    public static void makeClientOrderBlank() {

        HSSFWorkbook clientsBook = createBook(getClientsAndPricesFile());

        ArrayList<String> list = createHorListOfValues(clientsBook, "Цены", 4, 3);
        int clientNum = getNumberOfValueFromArray(list);
        ArrayList<String> price = createVertListOfValues(clientsBook, "Цены", 3, clientNum + 3, 0);

        closeBook(clientsBook);


        HSSFWorkbook blankBook = createBook(getBlankFile());
        writeVertListInBookSingle(blankBook, "Цены", 0, 3, price);
        writeBook(blankBook, getBlankFile());
        closeBook(blankBook);


        File file = new File(getBlankFile());
        // saveNewBlank();
        openFile(getBlankFile());

    }

    // IN - multiple or single client order, short directoryName
    // OUT - order in blank order
    // writed in same file, new list
    // OTHER: Write remains after order
    public static void createSortedOrder(String thisWorkDir) {

        /*
        ArrayList data = new ArrayList();
        data.add("some order.xls");
        data.add("order");
        data.add("order sorted");
        data.add(0);
        data.add(0);
        data.add(1);
        data.add(1);
        data.add(2);
*/

        ArrayList data = createDataForSorting(thisWorkDir);

        String orderFileName = data.get(0).toString();
        String sheetName = data.get(1).toString();
        String outSheetName = data.get(2).toString();
        int firstRowNum = (Integer) data.get(3);
        int artColNum = (Integer) data.get(4);
        int firstAmColNum = (Integer) data.get(5);
        int step = (Integer) data.get(6);
        int mask = (Integer) data.get(7);


        HSSFWorkbook orderBook = createBook(getWorkDir() + thisWorkDir + orderFileName);
        HashMap<String, ArrayList> map = createMapOrderMultiple(orderBook, sheetName, firstRowNum, artColNum, firstAmColNum, step);
        closeBook(orderBook);

        ArrayList<String> maskArray;

        if (mask == 1) {
            maskArray = createListOfArtsFromBlank();
        } else {
            maskArray = createListOfArtsFromPassport();
        }

        ArrayList orderList = createLongListOrderFromMultipleMap(map, maskArray);

        // Write it to out book
        HSSFWorkbook outBook = createBook(getWorkDir() + thisWorkDir + orderFileName);
        writeVertListInBookMultiple(outBook, outSheetName, 0, 0, 2, orderList);

        //Create List of remains in order of Passport
        //ArrayList<ArrayList> remainsArray = createShortListOrderFromSingleMap(map, createListOfArtsFromPassport());
        //writeRemainsListInBookMultiple(outbook, orderName, remainsArray);
        if (!map.isEmpty()) {
            printHashMap(map);
            ArrayList remainsList = createShortListOrderFromMultipleMap(map, createListOfArtsFromPassport());
            writeRemainsListInBookMultiple(outBook, outSheetName, remainsList);
        }

        writeBook(outBook, getWorkDir() + thisWorkDir + orderFileName);
        closeBook(outBook);

        openFile(getWorkDir() + thisWorkDir + orderFileName);

    }

    // IN - one file, short directoryName
    // You pick sheet, artNum etc
    // sheets must be identical by format!
    // OUT - if list of articuls in initial file
    // different than in actual passport -
    // rows will be past on theyr places
    public static void updateFileForNewPassport(String updateFileForNewPassportDir) {
        Scanner s = new Scanner(System.in);

        // Get a order file name
        ArrayList<String> dirAr = createDirectoryList(getWorkDir() + updateFileForNewPassportDir);

        System.out.println("Выбор файла для корректировки");
        String bookName = dirAr.get(chooseElementFromList(dirAr));

        HSSFWorkbook book = createBook(getWorkDir() + updateFileForNewPassportDir + bookName);

        // Get a sheet name
        ArrayList<String> sheetsAr = createSheetsList(book);
        System.out.println("\nВведите 1, если хотите скорректировать все листы, или 0, если какой-то конкретный");
        int choice = s.nextInt();

        //get a row num
        System.out.println("\nВведите номер строки, с которой начинается таблица заказа");
        int firstRowNum = s.nextInt();

        //get a art col
        System.out.println("\nВведите номер колонки с данными, содержащими артикул");
        int artColNum = s.nextInt();

        //create a actual doods list from passport
        ArrayList<ArrayList> actualPassportList = createListOfArtsAndNamesPassport();

        if (choice != 1) {
            System.out.println("\nВыбор листа для корректировки");
            String sheetName = sheetsAr.get(chooseElementFromList(sheetsAr));
            modifyListForActualPasport(book, sheetName, firstRowNum, artColNum, actualPassportList);
            closeBook(book);
        } else {
            for (String sheetName : sheetsAr) {
                modifyListForActualPasport(book, sheetName, firstRowNum, artColNum, actualPassportList);
            }
        }

        writeBook(book, getWorkDir() + updateFileForNewPassportDir + bookName);
        closeBook(book);

    }

    // IN - one file (our client order blank, old version), short directoryName
    // and other data (including date of planning shipment)
    // OUT - same order in special folder our actual blank
    // OTHER: Write remains after order
    public static void createSingleOrderFromOldOrderBlank(String thisWorkDir) {

        ArrayList data = createDataForNewOrder(thisWorkDir);

        /*
        ArrayList data = new ArrayList();
        data.add("Бланк заказа Александрова (ВТБ).xls");
        data.add("ФОРМА");
        data.add("ФОРМА sorted");
        data.add(9);
        data.add(0);
        data.add(4);
        data.add(1);
        data.add(1);

*/

        String orderFileName = data.get(0).toString();
        String sheetName = data.get(1).toString();
        String outSheetName = data.get(2).toString();
        int firstRowNum = (Integer) data.get(3);
        int artColNum = (Integer) data.get(4);
        int firstAmColNum = (Integer) data.get(5);

        Scanner s = new Scanner(System.in);
        System.out.println("На какую дату делаем заказ?");
        String date = s.next();


        //create order  in order of blank
        HSSFWorkbook book = createBook(getWorkDir() + thisWorkDir + orderFileName);
        HashMap<String, Integer> orderMap = createMapOrderSingle(book, sheetName, firstRowNum, artColNum, firstAmColNum);
        //String inn = String.valueOf(book.getSheet("Форма").getRow(3).getCell(1).getNumericCellValue());
        //int inn = (int) book.getSheet("Форма").getRow(3).getCell(1).getNumericCellValue();
        double inn = book.getSheet("Форма").getRow(3).getCell(1).getNumericCellValue();
        closeBook(book);

        // Definition of client, creating arrayList of prices
        HSSFWorkbook bookClients = createBook(getWorkDir() + "Клиенты и цены.xls");
        ArrayList innList = createHorListOfValues(bookClients, "Цены", 33, 3);

        int clientNum = innList.indexOf(inn);

        ArrayList<String> price = createVertListOfValues(bookClients, "Цены", 3, clientNum + 3, 0);
        closeBook(bookClients);

        //create client blank, writing prices and order
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


        //Create List of remains in order of Passport
        if (!orderMap.isEmpty()) {
            //printHashMap(orderMap);
            ArrayList remainsList = createShortListOrderFromSingleMap(orderMap, createListOfArtsFromPassport());
            writeRemainsListInBookSingle(outBook, outSheetName, remainsList);
        }

        outBook.getSheet(outSheetName).getRow(5).getCell(1).setCellValue(date);

        String customer = price.get(1);
        String newFolderName = date + " " + customer + "/";
        String newFileName = date + " " + customer + ".xls";
        String newFile = getWorkDir() + thisWorkDir + newFolderName + newFileName;

        File dir = new File(getWorkDir() + thisWorkDir + newFolderName);
        dir.mkdir();

        System.out.println("\nЗаказ записан в файл \"" + newFileName + "\" в одноименную папку в указанной рабочей директории.");

        writeBook(outBook, newFile);
        closeBook(outBook);

        openFile(newFile);

    }

    // IN - one file (client SINGLE order), short directoryName
    // and other data (including date of planning shipment)
    // OUT - same order in special folder our actual blank
    // OTHER: Write remains after order
    public static void createSingleOrderFromClientOrder(String thisWorkDir) {

        ArrayList data = createDataForOrderFromClientOrder(thisWorkDir);

/*
        ArrayList data = new ArrayList();
        data.add("tf order.xls");
        data.add("order");
        data.add("ФОРМА");
        data.add(0);
        data.add(0);
        data.add(1);
        */

        String orderFileName = data.get(0).toString();
        String sheetName = data.get(1).toString();
        String outSheetName = data.get(2).toString();
        int firstRowNum = (Integer) data.get(3);
        int artColNum = (Integer) data.get(4);
        int firstAmColNum = (Integer) data.get(5);

        // determine client
        Scanner s = new Scanner(System.in);
        System.out.println("Кто сделал заказ?");
        HSSFWorkbook clientsBook = createBook(getClientsAndPricesFile());
        ArrayList<String> list = createHorListOfValues(clientsBook, "Цены", 4, 3);
        int clientNum = getNumberOfValueFromArray(list);
        ArrayList<String> price = createVertListOfValues(clientsBook, "Цены", 3, clientNum + 3, 0);
        closeBook(clientsBook);

        System.out.println("\nУточняющая информация (например, какой магазин)?");
        String addInfo = s.next();

        System.out.println("\nНа какую дату делаем заказ?");
        String date = s.next();


        //create order  in order of blank
        HSSFWorkbook book = createBook(getWorkDir() + thisWorkDir + orderFileName);
        HashMap<String, Integer> orderMap = createMapOrderSingle(book, sheetName, firstRowNum, artColNum, firstAmColNum);
        closeBook(book);


        //create client blank, writing prices and order
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


        //Create List of remains in order of Passport
        if (!orderMap.isEmpty()) {
            //printHashMap(orderMap);
            ArrayList remainsList = createShortListOrderFromSingleMap(orderMap, createListOfArtsFromPassport());
            writeRemainsListInBookSingle(outBook, outSheetName, remainsList);
        }


        outBook.getSheet(outSheetName).getRow(5).getCell(1).setCellValue(date);
        outBook.getSheet(outSheetName).getRow(4).getCell(1).setCellValue(addInfo);

        String customer = price.get(1);
        String newFolderName = date + " " + customer + " " + addInfo + "/";
        String newFileName = date + " " + customer + " " + addInfo + ".xls";
        String newFile = getWorkDir() + thisWorkDir + newFolderName + newFileName;

        File dir = new File(getWorkDir() + thisWorkDir + newFolderName);
        dir.mkdir();

        System.out.println("\nЗаказ записан в файл \"" + newFileName + "\" в одноименную папку в указанной рабочей директории.");

        writeBook(outBook, newFile);
        closeBook(outBook);

        openFile(newFile);

    }

    // IN - one file (client MULTIPLE order), short directoryName
    // and other data (including date of planning shipment and row of addition data)
    // OUT - every order in special folder our actual blank
    // OTHER: Write remains after order (if you need this)
    public static void createMultipleOrderFromClientOrder(String thisWorkDir) {

/*
        ArrayList data = new ArrayList();
        data.add("ТФ шок заказ.xls");
        data.add("Общий");
        data.add("ФОРМА");
        data.add(1);
        data.add(0);
        data.add(1);
        data.add(1);
        data.add(1);

        */

        ArrayList data = createDataForMultipleOrder(thisWorkDir);

        // create data set for work
        String orderFileName = data.get(0).toString();
        String sheetName = data.get(1).toString();
        String outSheetName = data.get(2).toString();
        int firstRowNum = (Integer) data.get(3);
        int artColNum = (Integer) data.get(4);
        int firstAmColNum = (Integer) data.get(5);
        int step = (Integer) data.get(6);

        // get addInfo row
        Scanner s = new Scanner(System.in);
        System.out.println("\nНомер строки, в которой дополнительные данные для разнарядок?");
        int headRowNum = s.nextInt();


        //create multiple order map
        HSSFWorkbook orderBook = createBook(getWorkDir() + thisWorkDir + orderFileName);
        HashMap<String, ArrayList> orderMap = createMapOrderMultiple(orderBook, sheetName, firstRowNum, artColNum, firstAmColNum, step);

        //create Array of add information
        ArrayList<String> addInfoArray = createHorListOfValues(orderBook, sheetName, headRowNum, firstAmColNum);
        closeBook(orderBook);

        //find amount of oders, check it
        int am = 0;
        for (Map.Entry<String, ArrayList> entry : orderMap.entrySet()) {
            am = entry.getValue().size();
            if (am == 0) {
                continue;
            }
            //System.out.println("Amount: " + am);

            if (am != addInfoArray.size()) {
                System.out.println("Стоит проверить строку добавочной информации (количество не сопадает с нужным)");
                continue;
            }

            break;
        }

        // determine client
        System.out.println("Кто сделал заказ?");
        HSSFWorkbook clientsBook = createBook(getClientsAndPricesFile());
        ArrayList<String> list = createHorListOfValues(clientsBook, "Цены", 4, 3);
        int clientNum = getNumberOfValueFromArray(list);
        ArrayList<String> price = createVertListOfValues(clientsBook, "Цены", 3, clientNum + 3, 0);
        closeBook(clientsBook);

        // get a data of shipment
        System.out.println("\nНа какую дату делаем заказ?");
        String date = s.next();

        System.out.println("\nПечатать заказанные, но выведенные позиции? y/n");
        String remains = s.next();

        String customer = price.get(1);
        String newFolderName = date + " " + customer + "/";


        File dir = new File(getWorkDir() + thisWorkDir + newFolderName);
        dir.mkdir();


        for (int i = 0; i < am; i++) {


            String addInfo = String.valueOf(addInfoArray.get(i));
            HashMap<String, Integer> thisOrderMap = new HashMap<String, Integer>();

            // create this order map
            for (Map.Entry<String, ArrayList> entry : orderMap.entrySet()) {

                ArrayList<Integer> thisList = entry.getValue();

                if ((thisList.isEmpty()) || (i >= thisList.size())) {
                    continue;
                }

                int thisAm = thisList.get(i);

                //System.out.println(thisAm);
                if (thisAm == 0) {
                    continue;
                }

                thisOrderMap.put(entry.getKey(), thisAm);

            }


            //create client blank, writing prices and order
            HSSFWorkbook outBook = createBook(getBlankFile());
            writeVertListInBookSingle(outBook, "Цены", 0, 3, price);
            HSSFSheet sheet = outBook.getSheet(outSheetName);


            for (int e = 9; e < outBook.getSheet(outSheetName).getLastRowNum(); e++) {
                HSSFRow row = sheet.getRow(e);
                String art = row.getCell(0).getStringCellValue();
                if (thisOrderMap.containsKey(art)) {
                    row.getCell(4).setCellValue(thisOrderMap.get(art));
                    thisOrderMap.remove(art);
                } else {
                    continue;
                }
            }


            //Create List of remains in order of Passport
            if ((!thisOrderMap.isEmpty()) && (remains.equals("y"))) {
                //printHashMap(orderMap);
                ArrayList remainsList = createShortListOrderFromSingleMap(thisOrderMap, createListOfArtsFromPassport());
                writeRemainsListInBookSingle(outBook, outSheetName, remainsList);
            }


            outBook.getSheet(outSheetName).getRow(5).getCell(1).setCellValue(date);
            outBook.getSheet(outSheetName).getRow(4).getCell(1).setCellValue(addInfo);

            String newFileName = date + " " + customer + " " + addInfo + ".xls";
            String newFile = getWorkDir() + thisWorkDir + newFolderName + newFileName;

            System.out.println("\nЗаказ записан в файл \"" + newFileName + "\" в одноименную папку в указанной рабочей директории.");

            writeBook(outBook, newFile);
            closeBook(outBook);

        }

    }

    // Димону - перенести цифры из общих продаж в план
    // Если надо - поменять названия файла отчета по отгрузкам!!!
    public static void createSalesPlanForDimon() {

        HSSFWorkbook book1 = createBook(getWorkDir() + "Отчет отгрузки 2.xls");
        String sheetName1 = "Самый общий";
        HashMap<String, Integer> map = createMapOrderSingle(book1, sheetName1, 3, 1, 11);
        closeBook(book1);

        HSSFWorkbook book2 = createBook(getWorkDir() + "бланк План продаж.xls");
        HSSFSheet sheet = book2.getSheet("Лист11");


        for (int i = 3; i < 401; i++) {

            String articul = sheet.getRow(i).getCell(0).getStringCellValue();

            //System.out.println(i + ", " + articul);

            if (map.containsKey(articul)) {
                sheet.getRow(i).getCell(3).setCellValue(map.get(articul));
                //System.out.println("Yep! " + map.get(articul));
                map.remove(articul);
            } else {
                System.out.println("----------Nope! " + articul);
            }

        }

        writeBook(book2, getWorkDir() + "бланк План продаж.xls");
        closeBook(book2);

        openFile(getWorkDir() + "бланк План продаж.xls");


    }

    public static void createOneReportFromMultipleFiles() {

        //Create list
        //for every file  create map
        //
        //
    }

}

