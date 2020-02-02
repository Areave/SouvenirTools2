package ToolsFiles.library;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Array;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.io.*;

import static ToolsFiles.library.ShipmentStore.createShipment;
import static ToolsFiles.library.ShipmentStore.createShipmentForCashShipmentShort;
import static ToolsFiles.library.StoreUtil.*;
import static ToolsFiles.library.StoreWriting.*;
import static ToolsFiles.library.StoreCreating.*;

public class Test {

    static double addD = 0.0000000001;

    public static void main() {

        putPricesToReturnActs();

    }

    public static void putPricesToReturnActs() {
        //create map of prices
        HSSFWorkbook pricesBook = createBook(getWorkDir() + "temp/Себестоимость.xls");
        HSSFSheet sheet = pricesBook.getSheet("СС-продажа");

        HashMap<String, Double> pricesMap = new HashMap<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            //System.out.println(i);

            if ((sheet.getRow(i) == null) || (sheet.getRow(i).getCell(1) == null)) {
                continue;
            } else {

                String art = sheet.getRow(i).getCell(1).getStringCellValue();

                //String art = findArt(art0);


                double d = sheet.getRow(i).getCell(2).getNumericCellValue();

                // if (sheet.getRow(i).getCell(2).getNumericCellValue()) {continue;}

                DecimalFormat f = new DecimalFormat("#.##");
                double d3 = Double.parseDouble(f.format(d).replace(',', '.'));
                pricesMap.put(art, d3);
                closeBook(pricesBook);
            }
        }

        //printHashMap(pricesMap);

        //create list of files
        ArrayList<String> dirList = createDirectoryList(getWorkDir() + "temp/");
        dirList.remove("Себестоимость.xls");


        for (int u = 0; u < dirList.size(); u++) {

            String fileName = dirList.get(u);

            if (!fileName.equals("акт 341.xls")) {
                //dirList.remove(fileName);
                //u--;
            }

        }


        HashSet<String> set = new HashSet<>();

        // write prices in every file
        for (String fileName : dirList) {
            HSSFWorkbook book = createBook(getWorkDir() + "temp/" + fileName);
            HSSFSheet sheet2 = book.getSheet("TDSheet");
            for (int i = 9; i <= sheet2.getLastRowNum(); i++) {
                HSSFRow row = sheet2.getRow(i);
                if ((row == null) || (row.getCell(3)) == null) {
                    continue;
                }

                String art = row.getCell(3).getStringCellValue();

                if (art.contains(" ")) {
                    art = art.trim();
                }

                String art2;

                if (art == null) {
                    continue;
                }

                if (art.contains("p")) {
                    art2 = art.replace('p', 'р');
                } else if (art.contains("р")) {
                    art2 = art.replace('р', 'p');
                } else {
                    art2 = art;
                }

                if (pricesMap.containsKey(art)) {
                    row.createCell(40).setCellValue(pricesMap.get(art));
                    double sum = (pricesMap.get(art) * (row.getCell(24).getNumericCellValue()));
                    row.createCell(41).setCellValue(sum);
                } else if (pricesMap.containsKey(art2)) {
                    row.createCell(40).setCellValue(pricesMap.get(art2));
                    double sum = (pricesMap.get(art2) * (row.getCell(24).getNumericCellValue()));
                    row.createCell(41).setCellValue(sum);
                } else {
                    System.out.println("not contains! " + art + ", " + art2);
                    set.add(art);
                }
            }
            writeBook(book, getWorkDir() + "temp/" + fileName);
            closeBook(book);

        }

        //printSet(set);
        ArrayList<String> ar = new ArrayList<>();

        for (String s : set) {
            ar.add(s);
        }

        Collections.sort(ar);
        printArray(ar);
        //
        //
        //
    }

    public static String getRightNameForCashClient(String name, HashMap<String, String> nameMap) {

        String rightName;

        if (!nameMap.containsKey(name)) {
            System.out.println("Error! Нет имени " + name + " в файле клиентов и цен!");
            return null;
        } else {
            rightName = nameMap.get(name);
        }

        return rightName;
    }

    public static HashMap<String, String> createCommonNameForCashClientMap() {

        HashSet<String> set = createSetOfCashNamesFromClAndPr();
        HSSFWorkbook book = createBook(getClientsAndPricesFile());
        HSSFSheet sheet = book.getSheet("ДЖК");
        HashMap<String, String> cashNames = new HashMap<>();
        int l = sheet.getLastRowNum();

        for (int i = 0; i <= l; i++) {
            String name1 = sheet.getRow(i).getCell(0).getStringCellValue();
            if (name1.equals("-")) {
                continue;
            }
            String rightName = sheet.getRow(i).getCell(1).getStringCellValue();
            cashNames.put(name1, rightName);
        }

        return cashNames;
    }

    public static void threadMethod() {
        Thread ct = Thread.currentThread();
        System.out.println(ct);
        ct.setName("newThread");
        System.out.println(ct.getName());
        System.out.println(ct.getId());
        ct.setPriority(1);

        System.out.println(ct.getPriority());

        Thread nct = new Thread();
    }

    public static void checkDirOfShipmentsForErrors(String workDir) {
        ArrayList<String> l = new ArrayList();
        l.add("cashFiles/");
        l.add("bankFiles/");


        System.out.println("Проверка!");
        System.out.println();

        for (String s : l) {

            String thisWorkDir = workDir + s;
            ArrayList<String> dirList = createDirectoryList(thisWorkDir);
            for (String fileName : dirList) {
                checkShipmentFileForErrors(thisWorkDir, fileName);
            }
        }
    }

    public static void checkShipmentFileForErrors(String thisWorkDir, String fileName) {

        String fileNameFull = thisWorkDir + fileName;
        //if (!fileName.equals("10.01 Рудой — копия.xls")) {return;}
        System.out.println(fileName);


        // Check for format
        if (!fileName.endsWith("xls")) {
            System.out.println("Error! Неверный формат файла!");
            return;
        }

        HSSFWorkbook book = createBook(fileNameFull);
        HSSFSheet sheet = book.getSheet("ФОРМА");
        int last = sheet.getLastRowNum();


        // Check for date
        HSSFCell dateCell = sheet.getRow(5).getCell(1);
        if (dateCell == null) {
            System.out.println("Error! Нет даты!");
        }
        Date date = dateCell.getDateCellValue();
        if (date == null) {
            System.out.println("Error! Нет даты!");

        }

        // Check for normal summ
        int s = 0;
        double summ = createShipment(book).getSumm();
        last = sheet.getLastRowNum();

        for (int i = 11; i <= 300; i++) {

            if (i == 300) {
                System.out.println("Error! Что-то не так с суммой!");
            }

            HSSFRow row = sheet.getRow(i);
            if (row == null) {continue;}
            if (row.getCell(5) == null) {continue;}

            double b = row.getCell(5).getNumericCellValue();
            s += b;

            if ((b == summ)&&((s - b) == summ)) {
                break;
            }
        }

        System.out.println();

    }

    public static void createReportSummFile(String year, int fMonth, int lMonth) {

        String workDirString = getWorkDir() + "Shipment/" + year + "/";

        System.out.println("Год " + year);

        for (int i = fMonth; i <= lMonth; i++) {
            System.out.println("Месяц " + i);
            System.out.println();
            String workDirStringFull = workDirString + String.valueOf(i) + "/allFiles/";
            putShipmentShortOfMonthToReport(workDirStringFull, year, String.valueOf(i));
        }
    }

    public static void tempCreateAllFilesDirectories() {
        String workDirString = getWorkDir() + "Shipment/2019/";

        for (int i = 1; i <= 12; i++) {

            System.out.println("Месяц " + i);
            String workDirStringFull = workDirString + String.valueOf(i);

            File workDir = new File(workDirStringFull + "/allFiles");

            if (!workDir.exists()) {
                workDir.mkdir();
            } else {
                //  workDir.delete();
            }

            ArrayList<String> list = new ArrayList<String>();
            list.add(workDirStringFull + "/cashFiles/");
            list.add(workDirStringFull + "/bankFiles/");

            for (String dir : list) {

                ArrayList<String> list2 = createDirectoryList(dir);

                for (String fileName2 : list2) {

                    File fileFrom = new File(dir + fileName2);
                    File fileTo = new File(workDirStringFull + "/allFiles/" + fileName2);
                    copyFile(fileFrom, fileTo);
                }
            }

            // Check
            int sum = createDirectoryList(workDirStringFull + "/allFiles/").size();
            int ar1 = createDirectoryList(workDirStringFull + "/cashFiles/").size();
            int ar2 = createDirectoryList(workDirStringFull + "/bankFiles/").size();

            if (!(sum == (ar1 + ar2))) {
                System.out.println("Сумма не равна!");
                System.out.println(sum + ", " + ar1 + ", " + ar2);
            } else {
                System.out.println("Все ок!");
            }
        }
    }

    public static void createAllFilesDirectoryForMonth(String workDirStringFull) {

        File workDir = new File(workDirStringFull + "/allFiles");

        if (!workDir.exists()) {
            workDir.mkdir();
        } else {
            workDir.delete();
        }

        ArrayList<String> list = new ArrayList<String>();
        list.add(workDirStringFull + "/cashFiles/");
        list.add(workDirStringFull + "/bankFiles/");

        for (String dir : list) {

            ArrayList<String> list2 = createDirectoryList(dir);

            for (String fileName2 : list2) {

                File fileFrom = new File(dir + fileName2);
                File fileTo = new File(workDirStringFull + "/allFiles/" + fileName2);
                copyFile(fileFrom, fileTo);
            }
        }

        // Check
        int sum = createDirectoryList(workDirStringFull + "/allFiles/").size();
        int ar1 = createDirectoryList(workDirStringFull + "/cashFiles/").size();
        int ar2 = createDirectoryList(workDirStringFull + "/bankFiles/").size();

        if (!(sum == (ar1 + ar2))) {
            System.out.println("Сумма не равна!");
            System.out.println(sum + ", " + ar1 + ", " + ar2);
        } else {
            System.out.println("Все ок!");
        }
    }


    public static void putShipmentShortOfMonthToReport(String workDir, String year, String month) {


        // Create List of fileNames
        ArrayList<String> dirList = createDirectoryList(workDir);

        // Create Map for shipmentsShort
        HashMap<Long, ArrayList<ShipmentShort>> map = new HashMap<Long, ArrayList<ShipmentShort>>();
        // for every file Name create shipmentShort

        for (String fileName : dirList) {
            ArrayList shipList = new ArrayList();

            HSSFWorkbook book = createBook(workDir + fileName);
            Shipment shipment = createShipment(book);

            ShipmentShort shipmentShort;

            if (shipment.getInn() == 7706092528l) {
                shipment = createShipmentForCashShipmentShort(book);
                shipmentShort = new ShipmentShort(shipment, shipment.getCustomer());
            } else {
                shipmentShort = new ShipmentShort(shipment);
            }

            Long inn = shipmentShort.getInn();

            ArrayList list;

            if (map.containsKey(inn)) {
                list = map.get(inn);
            } else {
                list = new ArrayList();
            }

            // put it in map
            list.add(shipmentShort);
            Collections.sort(list);
            map.put(inn, list);

        }

        // Now we have a map with inn as a key and array of shipmentShort
        // as a value for this key

        // Create book and sheetName
        HSSFWorkbook reportBook = createBook(getWorkDir() + "Отчет отгрузки по суммам.xls");

        String sheetName = year.substring(2, 4) + "_" + month;

        if (!(reportBook.getSheet(sheetName) == null)) {
            System.out.println("Sheet is exist!");
            sheetName = sheetName + "_1";
        }

        HSSFSheet sheet = reportBook.createSheet(sheetName);


        // Write data in sheet
        // Cash First
        Long cashInn = 7706092528l;

        if (map.containsKey(cashInn)) {
            ArrayList<ShipmentShort> shipmentShortArray = map.get(cashInn);
            writeShipmentShortArrayInBook(sheet, shipmentShortArray);
            map.remove(cashInn);
        }

        //Writing rest of shipments
        for (Map.Entry<Long, ArrayList<ShipmentShort>> entry : map.entrySet()) {
            ArrayList list = entry.getValue();
            writeShipmentShortArrayInBook(sheet, list);
        }

        sheet.setColumnWidth(0, 0 * 256);
        sheet.setColumnWidth(1, 15 * 256);
        sheet.setColumnWidth(3, 19 * 256);
        sheet.setColumnWidth(4, 10 * 256);
        sheet.setColumnHidden(2, true);

        writeBook(reportBook, getWorkDir() + "Отчет отгрузки по суммам.xls");
        closeBook(reportBook);
        //openFile(getWorkDir() + "Отчет отгрузки по суммам.xls");

        //
        //
        //

    }

    public static void writeShipmentShortArrayInBook(HSSFSheet sheet, ArrayList<ShipmentShort> shipmentShortArray) {

        //create styles
        HSSFCellStyle styleData = sheet.getWorkbook().createCellStyle();
        styleData.setAlignment(HorizontalAlignment.LEFT);


        HSSFCellStyle styleFormula = sheet.getWorkbook().createCellStyle();
        HSSFFont formulaFont = sheet.getWorkbook().createFont();
        formulaFont.setBold(true);
        formulaFont.setFontHeight((short) 200);
        styleFormula.setFont(formulaFont);


        if ((sheet.getRow(0) == null)) {
            sheet.createRow(0).createCell(0).setCellValue("-");
            sheet.getRow(0).createCell(1).setCellValue("Клиент");
            sheet.getRow(0).createCell(2).setCellValue("ИНН");
            sheet.getRow(0).createCell(3).setCellValue("Дата отгрузки");
            sheet.getRow(0).createCell(4).setCellValue("Сумма");
            sheet.getRow(0).createCell(5).setCellValue("Итого");

            HSSFFont headerFont = sheet.getWorkbook().createFont();
            headerFont.setBold(true);
            headerFont.setFontHeight((short) 200);

            CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
            styleHeader.setFont(headerFont);
            styleHeader.setFillBackgroundColor(IndexedColors.DARK_YELLOW.getIndex());
            styleHeader.setAlignment(HorizontalAlignment.LEFT);
            styleHeader.setBorderLeft(BorderStyle.THIN);
            styleHeader.setBorderTop(BorderStyle.THIN);
            styleHeader.setBorderRight(BorderStyle.THIN);
            styleHeader.setBorderBottom(BorderStyle.THIN);


            for (int n = 1; n <= 5; n++) {
                sheet.getRow(0).getCell(n).setCellStyle(styleHeader);
            }


        }

        int last = sheet.getLastRowNum();
        int value = shipmentShortArray.size();

        HashMap<Long, String> innMap = createMapOfInnFromClAndPr();

        String start, finish;
        double summa = 0;

        for (int i = 0; i < value; i++) {

            ShipmentShort shipmentShort = shipmentShortArray.get(i);
            int row = i + last + 1;

            long inn = shipmentShort.getInn();
            String name;

            if (inn == 7706092528l) {
                name = shipmentShort.getCustomer();
            } else {
                name = getNameOfClientByInnWithInnMap(shipmentShort.getInn(), innMap);
            }

            String date = shipmentShort.createSimpleDate(shipmentShort.getShipDate());
            Date dateD = shipmentShort.getShipDate();
            Double summ = shipmentShort.getSumm();

            HSSFCell cellD = sheet.createRow(row).createCell(0);
            cellD.setCellValue("-");

            HSSFCell cellName = sheet.getRow(row).createCell(1);
            cellName.setCellValue(name);

            HSSFCell cellINN = sheet.getRow(row).createCell(2);
            cellINN.setCellValue(inn);

            HSSFCell cellDate = sheet.getRow(row).createCell(3);
            //cellDate.setCellType(CellType.NUMERIC);
            cellDate.setCellValue(date);

            HSSFCell cellSumm = sheet.getRow(row).createCell(4);
            cellSumm.setCellValue(summ);
            summa = summa + summ;

            for (int n = 1; n <= 4; n++) {
                sheet.getRow(row).getCell(n).setCellStyle(styleData);
            }
        }


        sheet.createRow(sheet.getLastRowNum() + 1).createCell(0).setCellValue("-");

        // put formula summ
        HSSFCell formulaCell = sheet.getRow(sheet.getLastRowNum() - 1).createCell(5);
        formulaCell.setCellStyle(styleFormula);
        formulaCell.setCellValue(summa);

        //not working correctly
        //start = sheet.getRow(sheet.getLastRowNum()-value + 1).getCell(4).getAddress().toString();
        //finish = sheet.getRow(sheet.getLastRowNum()).getCell(4).getAddress().toString();
        //formulaCell.setCellFormula("СУММ(" + start +":" + finish +")");


    }

    public static void comparasingShipment() {

        Scanner s = new Scanner(System.in);

        // Take a year for first part

        System.out.println("Сравнение отгрузок по периодам");
        System.out.println("Период в сколько месяцев будем сравнивать?");
        int per = s.nextInt();

        System.out.println("Введите год первого периода");
        int fYear = s.nextInt();

        int fMonth, lMonth;

        do {
            System.out.println("Введите первый месяц первого периода");
            fMonth = s.nextInt();
            lMonth = fMonth + per - 1;
            System.out.println(lMonth);
        }
        while (lMonth > 12);


        System.out.println("Введите год второго периода");
        int sYear = s.nextInt();

        System.out.println("Будут сравниваться: ");
        System.out.println(fYear + " год, месяцы " + fMonth + "-" + lMonth);
        System.out.println("и");
        System.out.println(sYear + " год, месяцы " + fMonth + "-" + lMonth);

        ArrayList<String> monthList = new ArrayList<>();
        for (int i = fMonth; i <= lMonth; i++) {
            monthList.add(String.valueOf(i));
        }


        // Create HashMap of shipments for first part

        HSSFWorkbook reportBook = createBook(getWorkDir() + "Отчет отгрузки по суммам.xls");

        for (int e = 0; e < monthList.size(); e++) {
            String sheetName = String.valueOf(fYear) + "_" + monthList.get(e);
            HSSFSheet sheet = reportBook.getSheet(sheetName);

            if (sheet == null) {
                System.out.println("Данных для " + monthList.get(e) + " месяца " + String.valueOf(fYear) + " года нет!");
                return;
            } else {
                for (int r = 1; r < sheet.getLastRowNum(); r++) {
                    HSSFRow row = sheet.getRow(r);
                    Long inn = (long) row.getCell(2).getNumericCellValue();
                    String dateString = row.getCell(3).getStringCellValue();
                    Date date = row.getCell(3).getDateCellValue();
                    double sum = row.getCell(4).getNumericCellValue();
                    ShipmentShort shipmentShort = new ShipmentShort(inn, date, sum);
                    System.out.println(shipmentShort);
                }
            }
        }


        // Create HashMap of shipments for second part
        // Find common clients for both part, put it in array(?)

        // Create book and sheet
        //HSSFWorkbook reportBook = createBook(getWorkDir() + "Отчет отгрузки по суммам.xls");

        // Take INN from this array and write in report
        //HashMap<Long, String> innMap = createMapOfInnFromClAndPr();

        // Take rest of INN and write it too


    }

    public static void comparasingShipment(String fYear, String sYear, int fMonth, int per) {


        // Take a year for first part

        System.out.println("Сравнение отгрузок по периодам");

        int lMonth = fMonth + per - 1;

        System.out.println("Будут сравниваться: ");
        System.out.println(fYear + " год, месяцы " + fMonth + "-" + lMonth);
        System.out.println("и");
        System.out.println(sYear + " год, месяцы " + fMonth + "-" + lMonth);
        System.out.println();


        ArrayList<String> monthList = new ArrayList<>();
        for (int i = fMonth; i <= lMonth; i++) {
            monthList.add(String.valueOf(i));
        }

        HSSFWorkbook reportBook = createBook(getWorkDir() + "Отчет отгрузки по суммам.xls");


        // Create HashMap of shipments for both years
        HashMap<String, ArrayList<ShipmentShort>> shipmentShortMapFirst = createMapForComparassing(reportBook, fYear, monthList);
        HashMap<String, ArrayList<ShipmentShort>> shipmentShortMapSecond = createMapForComparassing(reportBook, sYear, monthList);

        // Create comparassing and write it
        String comparassingSheetname = "Сравнение " + fYear.substring(2) + "-" + sYear.substring(2) + "год " + fMonth + "-" + lMonth + " мес";
        reportBook.createSheet(comparassingSheetname);

        HashMap<String, ArrayList<ArrayList>> commonShipmentMap = new HashMap<String, ArrayList<ArrayList>>();

        for (Map.Entry<String, ArrayList<ShipmentShort>> entry : shipmentShortMapFirst.entrySet()) {

            String name = entry.getKey();

            if (shipmentShortMapSecond.containsKey(name)) {
                ArrayList<ArrayList> l = new ArrayList<>();
                l.add(entry.getValue());
                l.add(shipmentShortMapSecond.get(name));
                commonShipmentMap.put(name, l);
            }
        }

        //Delete common shipments from both map
        for (Map.Entry<String, ArrayList<ArrayList>> entry : commonShipmentMap.entrySet()) {
            String name = entry.getKey();
            shipmentShortMapFirst.remove(name);
            shipmentShortMapSecond.remove(name);
        }

        //printHashMap(commonINNMap);

        HSSFSheet sheet = reportBook.getSheet(comparassingSheetname);

        sheet.createRow(0).createCell(0).setCellValue("Сравнение отгрузок за период " + fMonth + "-" + lMonth + " мес " + fYear + " и " + sYear + " годов");
        sheet.createRow(1).createCell(1).setCellValue(fYear + "г.");
        sheet.getRow(1).createCell(5).setCellValue(sYear + "г.");
        HSSFRow row2 = sheet.createRow(2);

        for (int i = 1; i < 6; i += 4) {

            row2.createCell(i).setCellValue("Клиент");
            row2.createCell(i + 1).setCellValue("Дата");
            row2.createCell(i + 2).setCellValue("Сумма");
        }

        //sheet.setColumnWidth(1, 15 * 256);


        writeComparrasingMapToSheet(sheet, commonShipmentMap);
        writeComparrasingDifferentMapToSheet(sheet, shipmentShortMapFirst, shipmentShortMapSecond);
        writeBook(reportBook, getWorkDir() + "Отчет отгрузки по суммам.xls");
        closeBook(reportBook);
        openFile(getWorkDir() + "Отчет отгрузки по суммам.xls");

    }

    public static void writeComparrasingDifferentMapToSheet(HSSFSheet sheet, HashMap<String, ArrayList<ShipmentShort>> map1, HashMap<String, ArrayList<ShipmentShort>> map2) {

        HashMap<String, Double> firstDifMap = new HashMap<String, Double>();
        HashMap<String, Double> secDifMap = new HashMap<String, Double>();
        int firstRowS = sheet.getLastRowNum() + 1;
        int firstRowS2 = sheet.getLastRowNum() + 1;
        System.out.println();

        //First
        for (Map.Entry<String, ArrayList<ShipmentShort>> entry : map1.entrySet()) {
            String name = entry.getKey();
            ArrayList<ShipmentShort> ar = entry.getValue();
            double sum1 = 0;

            int firstRowF = sheet.getLastRowNum() + 1;

            for (ShipmentShort ship1 : ar) {

                HSSFRow row = sheet.createRow(firstRowF);
                writeShipmentShortToRow(name, ship1, row, 1);
                sum1 += ship1.getSumm();
                System.out.println("Write " + name + " in row firstRowF " + firstRowF);
                firstRowF++;

            }

            HSSFRow row = sheet.createRow(firstRowF);
            for (int rn = 0; rn <= 4; rn++) {
                row.createCell(rn).setCellValue("-");
            }

            firstDifMap.put(name, sum1);

        }


        //Second

        for (Map.Entry<String, ArrayList<ShipmentShort>> entry2 : map2.entrySet()) {
            String name2 = entry2.getKey();
            ArrayList<ShipmentShort> ar2 = entry2.getValue();
            double sum2 = 0;

            System.out.println(name2);
            System.out.println("firstRowS " + firstRowS);

            for (ShipmentShort ship2 : ar2) {

                HSSFRow row2 = sheet.getRow(firstRowS);

                if (row2 == null) {
                    row2 = sheet.createRow(firstRowS);
                    // System.out.println("sheet is null! create!!! " + firstRowS);
                }

                writeShipmentShortToRow(name2, ship2, row2, 5);

                sum2 += ship2.getSumm();
                System.out.println("Writed " + name2 + " in row " + firstRowS);
                firstRowS++;

            }


            HSSFRow row2div = sheet.getRow(firstRowS);
            System.out.println(firstRowS + " NOW-------");


            if (row2div == null) {
                row2div = sheet.createRow(firstRowS);
            }
            for (int rn = 5; rn <= 8; rn++) {
                row2div.createCell(rn).setCellValue("-");
            }

            firstRowS++;

            secDifMap.put(name2, sum2);

        }


        HSSFRow row = sheet.getRow(firstRowS2 + 1);
        if (row == null) {
            row = sheet.createRow(firstRowS2);
        }
        row.createCell(10).setCellValue("Появились клиенты:");

        HSSFRow row2 = sheet.getRow(firstRowS2 + 2);
        if (row2 == null) {
            row2 = sheet.createRow(firstRowS2);
        }
        row2.createCell(10).setCellValue("Клиент:");
        row2.createCell(11).setCellValue("Сумма отгр.:");

        int rowNewF = firstRowS2 + 3;
        int rowNewF2 = firstRowS2 + 3;

        for (Map.Entry<String, Double> entry : secDifMap.entrySet()) {
            String name = entry.getKey();
            Double sum = entry.getValue();

            HSSFRow row3 = sheet.getRow(rowNewF);
            if (row3 == null) {
                row3 = sheet.createRow(rowNewF);
            }

            row3.createCell(10).setCellValue(name);
            row3.createCell(11).setCellValue(sum);

            rowNewF++;

        }


        HSSFRow row0 = sheet.getRow(firstRowS2 + 1);
        if (row0 == null) {
            row0 = sheet.createRow(firstRowS2);
        }
        row0.createCell(13).setCellValue("Ушли клиенты:");

        HSSFRow row02 = sheet.getRow(firstRowS2 + 2);
        if (row02 == null) {
            row02 = sheet.createRow(firstRowS2);
        }
        row02.createCell(13).setCellValue("Клиент:");
        row02.createCell(14).setCellValue("Сумма отгр.:");


        for (Map.Entry<String, Double> entry : firstDifMap.entrySet()) {
            String name = entry.getKey();
            Double sum = entry.getValue();

            HSSFRow row3 = sheet.getRow(rowNewF2);
            if (row3 == null) {
                row3 = sheet.createRow(rowNewF2);
            }

            row3.createCell(13).setCellValue(name);
            row3.createCell(14).setCellValue(sum);

            rowNewF2++;

        }
    }


    public static void writeShipmentShortToRow(String name, ShipmentShort ship, HSSFRow row, int firstColNum) {
        row.createCell(firstColNum).setCellValue(name);
        row.createCell(firstColNum + 1).setCellValue(ship.createSimpleDate(ship.getShipDate()));
        row.createCell(firstColNum + 2).setCellValue(ship.getSumm());
    }

    public static void writeComparrasingMapToSheet(HSSFSheet sheet, HashMap<String, ArrayList<ArrayList>> commonShipmentMap) {

        int counter = 1;
        HashMap<String, Double> compMap = new HashMap<>();
        HashMap<String, Double> dinMap = new HashMap<>();

        for (Map.Entry<String, ArrayList<ArrayList>> entry : commonShipmentMap.entrySet()) {

            int size = commonShipmentMap.size();
            String name = entry.getKey();
            ArrayList<ShipmentShort> list1 = entry.getValue().get(0);
            ArrayList<ShipmentShort> list2 = entry.getValue().get(1);
            int fRow = sheet.getLastRowNum() + 1;
            int fRowForAn = sheet.getLastRowNum() + 1;
            double sum1 = 0;
            double sum2 = 0;

            for (ShipmentShort ship1 : list1) {
                int rNum = sheet.getLastRowNum() + 1;
                sheet.createRow(rNum).createCell(1).setCellValue(name);
                sheet.getRow(rNum).createCell(2).setCellValue(ship1.createSimpleDate(ship1.getShipDate()));
                sheet.getRow(rNum).createCell(3).setCellValue(ship1.getSumm());
                sum1 += ship1.getSumm();
            }

            for (ShipmentShort ship2 : list2) {
                HSSFRow rowLast = sheet.getRow(fRow);
                if (rowLast == null) {
                    rowLast = sheet.createRow(fRow);
                }
                rowLast.createCell(5).setCellValue(name);
                rowLast.createCell(6).setCellValue(ship2.createSimpleDate(ship2.getShipDate()));
                rowLast.createCell(7).setCellValue(ship2.getSumm());
                sum2 += ship2.getSumm();
                fRow++;
            }


            HSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);

            for (int rn = 0; rn < 8; rn++) {
                row.createCell(rn).setCellValue("-");
            }

            double dif = sum2 - sum1;
            double percDinamic = (dif / sum1);

            System.out.println(name);
            System.out.println(dif);
            System.out.println(percDinamic);

            compMap.put(name, dif);
            dinMap.put(name, percDinamic);

            //Analisys

            /*

            sheet.getRow(fRowForAn).createCell(9).setCellValue("Итого за 1пер");
            sheet.getRow(fRowForAn).createCell(10).setCellValue("Итого за 2пер");
            HSSFRow rowAn = sheet.getRow(fRowForAn + 1);
            if (rowAn == null) {
                rowAn = sheet.createRow(fRow + 1);
            }
            rowAn.createCell(9).setCellValue(sum1);
            rowAn.createCell(10).setCellValue(sum2);



            String res;

            if (dif >= 0) {
                res = " увеличился";
            } else {
                res = " уменьшился";
            }

            sheet.getRow(fRowForAn).createCell(12).setCellValue("Объем отгрузок по " + name + res + " на");
            rowAn.createCell(12).setCellValue(dif);

            */


            if (counter == size) {
                System.out.println("===");
                sheet.getRow(1).createCell(10).setCellValue("Клиент");
                sheet.getRow(1).createCell(11).setCellValue("Динамика");
                sheet.getRow(1).createCell(13).setCellValue("В процентах");
                int row2 = 2;
                //printHashMap(compMap);

                for (Map.Entry<String, Double> entry3 : compMap.entrySet()) {
                    double dif2 = entry3.getValue();

                    sheet.getRow(row2).createCell(10).setCellValue(entry3.getKey());
                    sheet.getRow(row2).createCell(11).setCellValue(dif2);

                    if (dif2 < 0) {
                        sheet.getRow(row2).createCell(12).setCellValue("Падение!");
                    } else {
                        sheet.getRow(row2).createCell(12).setCellValue("Рост!");
                    }
                    double din = dinMap.get(entry3.getKey());
                    int intDin = (int) (100 * din);
                    sheet.getRow(row2).createCell(13).setCellValue(intDin);
                    System.out.println(entry3.getKey() + ", " + intDin);

                    row2++;

                }
            }

            counter++;

        }

    }

    public static HashMap<String, ArrayList<ShipmentShort>> createMapForComparassing(HSSFWorkbook reportBook, String year, ArrayList<String> monthList) {

        HashMap<String, ArrayList<ShipmentShort>> shipmentShortMap = new HashMap<String, ArrayList<ShipmentShort>>();

        for (int e = 0; e < monthList.size(); e++) {
            String sheetName = String.valueOf(year).substring(2) + "_" + monthList.get(e);
            HSSFSheet sheet = reportBook.getSheet(sheetName);
            HashMap innMap = createMapOfInnFromClAndPr();

            if (sheet == null) {
                System.out.println("Данных для " + monthList.get(e) + " месяца " + String.valueOf(year) + " года нет!");
                return null;
            } else {
                for (int r = 1; r < sheet.getLastRowNum(); r++) {
                    HSSFRow row = sheet.getRow(r);

                    if (row.getCell(2) == null) {
                        continue;
                    }

                    Long inn = (long) row.getCell(2).getNumericCellValue();
                    String name = row.getCell(1).getStringCellValue();

                    //String name = getNameOfClientByInnWithInnMap(inn, innMap);
                    String dateString = row.getCell(3).getStringCellValue();

                    SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                    String dataString = row.getCell(3).getStringCellValue();
                    Date date = null;
                    try {
                        date = format.parse(dataString);
                    } catch (ParseException ex) {
                        ex.getMessage();
                        System.out.println("Дата не спарсилась" + dataString);
                    }

                    double sum = row.getCell(4).getNumericCellValue();
                    ShipmentShort shipmentShort = new ShipmentShort(inn, date, sum);

                    ArrayList<ShipmentShort> array;

                    if (shipmentShortMap.containsKey(name)) {
                        array = shipmentShortMap.get(name);

                    } else {
                        array = new ArrayList<ShipmentShort>();
                    }

                    array.add(shipmentShort);
                    shipmentShortMap.put(name, array);
                }
            }
        }

        return shipmentShortMap;
    }

    public static double getSummOfShipmentShortFromMap(HashMap<String, ArrayList<ShipmentShort>> shipmentShortMap) {

        double sum = 0;

        for (HashMap.Entry<String, ArrayList<ShipmentShort>> entry : shipmentShortMap.entrySet()) {
            ArrayList<ShipmentShort> array = entry.getValue();
            for (ShipmentShort shipment : array) {
                sum += shipment.getSumm();
            }
        }

        return sum;
    }

    // Take a inn long as a parameter
    // return name of client
    // taken from inn map
    // from ClAndPr sheet "ИНН"
    // lso take a innMap as a parameter
    // for usability
    public static String getNameOfClientByInn(long inn) {

        String name;

        HashMap<Long, String> innMap = createMapOfInnFromClAndPr();

        if (innMap.containsKey(inn)) {
            name = innMap.get(inn);
            return name;
        } else {
            return null;
        }

    }

    public static String getNameOfClientByInnWithInnMap(long inn, HashMap<Long, String> innMap) {

        String name;

        if (innMap.containsKey(inn)) {
            name = innMap.get(inn);
            return name;
        } else {
            return null;
        }

    }


    // Take a directory as an argument
    // Check every ship from there, take a name and INN
    // Create a mapInn from clientsAndPrices sheet "ИНН"
    // if there is no exactly that INN - put it in to ClAndPr to this sheet
    public static void putInnFromEveryShipsInDirToClAndPrices(String workDirShip) {

        ArrayList<String> dirList = createDirectoryList(workDirShip);

        for (String fileName : dirList) {

            HashMap<Long, String> innMap = createMapOfInnFromClAndPr();
            System.out.println("INN Map was created, it's " + innMap.size());

            HSSFWorkbook ship = createBook(workDirShip + fileName);
            HSSFSheet sheet = ship.getSheet("Форма");
            HSSFCell innCell = sheet.getRow(3).getCell(1);
            long inn = (long) innCell.getNumericCellValue();
            String name = sheet.getRow(2).getCell(1).getStringCellValue();
            closeBook(ship);

            System.out.println(name + ", " + inn);

            if (!innMap.containsKey(inn)) {

                System.out.println("Add " + name + ", " + inn);


                HSSFWorkbook clients = createBook(getClientsAndPricesFile());
                HSSFSheet innSheet = clients.getSheet("ИНН");
                int lastRow = innSheet.getLastRowNum();
                System.out.println("Last row still is " + lastRow);

                innSheet.createRow(lastRow + 1).createCell(1).setCellValue(inn);
                innSheet.getRow(lastRow + 1).createCell(0).setCellValue(name);
                writeBook(clients, getClientsAndPricesFile());
                closeBook(clients);

            }

        }

    }

    public static void tempPutNameOfClientFromEveryShipsInDirToClAndPrices(String year, int month) {

        String workDirShip = getWorkDir() + "Shipment/" + year + "/" + String.valueOf(month) + "/allFiles/";
        ArrayList<String> dirList = createDirectoryList(workDirShip);

        for (String fileName : dirList) {

            //HashMap<Long, String> cashNameMap = createMapOfInnFromClAndPr();

            HSSFWorkbook ship = createBook(workDirShip + fileName);
            HSSFSheet sheet = ship.getSheet("Форма");
            HSSFCell innCell = sheet.getRow(3).getCell(1);
            long inn = (long) innCell.getNumericCellValue();
            String name = sheet.getRow(2).getCell(1).getStringCellValue();
            closeBook(ship);

            if (!(inn == 7706092528l)) {
                continue;
            } else {

                System.out.println("Add " + name + ", " + inn);
                HSSFWorkbook clients = createBook(getClientsAndPricesFile());
                HSSFSheet nameSheet = clients.getSheet("ДЖК");
                int lastRow = nameSheet.getLastRowNum();
                System.out.println("Last row still is " + lastRow);
                nameSheet.createRow(lastRow + 1).createCell(0).setCellValue(name);
                writeBook(clients, getClientsAndPricesFile());
                closeBook(clients);

            }

        }

    }

    //take a number of month,
    // print every sheet name,
    // that have no digits in summ sales cell
    // from january
    public static void findClientsWithEmptyReports(int m) {

        HSSFWorkbook book = createBook(getWorkDir() + "CommonSalesReport/Отчет продажи.xls");

        ArrayList<String> list = createHorListOfValues(book, "Общий отчет", 0, 2);

        //int n = chooseElementFromList(list);

        //System.out.println(n);

        ArrayList<String> sheetList = createSheetsList(book);

        for (int n = m - 1; n < m; n++) {

            String month = list.get(n);

            ArrayList<String> noFillSheetList = new ArrayList<String>();

            for (String s : sheetList) {

                if (s.equals("Чистый бланк")) {
                    continue;
                }

                HSSFSheet sheet = book.getSheet(s);
                //System.out.println(s);
                HSSFCell cell = sheet.getRow(1338).getCell((2 * n) + 3);

                if ((cell == null) || (cell.getNumericCellValue() == 0.0)) {

                    noFillSheetList.add(s);
                }

            }

            if (!noFillSheetList.isEmpty()) {
                noFillSheetList.remove("Общий отчет");
                System.out.println("Month :" + month + ", clients with empty data:");

                printArray(noFillSheetList);
                noFillSheetList.clear();
            }

        }

    }

    //сверяет акты, выводит в консоль то, что не сходится
    public static void sverka() {

        HSSFWorkbook book = createBook(getWorkDir() + "Акт Караханов общий.xls");

        HSSFSheet karSheet = book.getSheet("kar");
        HSSFSheet proSheet = book.getSheet("pro");

        HashSet<Double> setKar = createSetOfPaymentForActComparison(book, karSheet, 1);
        HashSet<Double> setPro = createSetOfPaymentForActComparison(book, proSheet, 1);

        HashSet<Double> setKar2 = (HashSet<Double>) setKar.clone();
        HashSet<Double> setPro2 = (HashSet<Double>) setPro.clone();

        /*
        System.out.println("Что есть в их акте");
        printSet(setKar);
        System.out.println("Что есть в нашем акте");
        printSet(setPro);
        */

        Iterator<Double> iter = setKar.iterator();
        while (iter.hasNext()) {
            double b = iter.next();
            if (setPro.contains(b)) {
                setKar2.remove(b);
                setPro2.remove(b);
            }
        }

        System.out.println("Что осталось в их акте");
        printSet(setKar2);
        System.out.println("Что осталось в нашем акте");
        printSet(setPro2);

    }

    public static HashSet createSetOfPaymentForActComparison(HSSFWorkbook book, HSSFSheet sheet, int columnNum) {

        HashSet set = new HashSet();

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            HSSFCell cell = sheet.getRow(i).getCell(columnNum);

            if (cell == null) {
                continue;
            }

            double d = cell.getNumericCellValue();

            if (d == 0.0) {
                continue;
            }

            if (set.contains(d)) {
                System.out.println("--------Repeat!!!! " + d + ", row is " + i);
                d = d + addD;
            }

            set.add(d);

        }

        return set;
    }

    //перехуячивает димонов план продаж
    public static void main2() {

        String dir = getWorkDir();
        String book1name = "Отчет отгрузки.xls";
        String book1file = dir + book1name;
        String book2name = "План продаж.xls";
        String book2file = dir + book2name;

        HSSFWorkbook book1 = createBook(book1file);
        HashMap<String, Integer> map = createMapOrderSingle(book1, "Самый общий", 3, 1, 10);
        closeBook(book1);

        //printHashMap(map);
        //System.out.println(mapValueSumm(map));

        HSSFWorkbook book2 = createBook(book2file);
        HSSFSheet sheet = book2.getSheet("sheet");

        for (int i = 3; i < 371; i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {

                System.out.println("null! " + i);

            }

            HSSFCell artCell = row.getCell(0);

            if (artCell == null) {
                continue;
            }

            String art = artCell.getStringCellValue();


            //System.out.println(art);

            if (map.containsKey(art)) {

                if (map.get(art) == 0) {
                    continue;
                }

                row.createCell(4).setCellValue(map.get(art));
                map.remove(art);

            }
        }


        if (!map.isEmpty()) {
            printHashMap(map);
            ArrayList list = createShortListOrderFromSingleMap(map, createListOfArtsFromPassport());
            writeRemainsListInBookSingle(book2, "sheet", list);
        }

        writeBook(book2, book2file);
        closeBook(book2);
        openFile(book2file);
    }

    public static void printSet(HashSet set) {

        Iterator it = set.iterator();
        while (it.hasNext()) {
            System.out.println(it.next());
        }
    }

    public static void main3() {

        String thisWorkDirMag = getWorkDir() + getTfReportsDir() + "mag/";

        ArrayList<String> dirAr = createDirectoryList(thisWorkDirMag);

        for (String filename : dirAr) {

            System.out.println(filename);

            HSSFWorkbook book = createBook(thisWorkDirMag + filename);
            HashMap map = createMapOrderSingle(book, book.getSheetAt(0).getSheetName(), 0, 0, 1);
            closeBook(book);

            ArrayList ar = createLongListOrderFromSingleMap(map, createListOfArtsFromPassport());

            HSSFWorkbook book2 = createBook(thisWorkDirMag + "tfMagRep.xls");
            writeArrayArrayStringInt(book2, filename, 0, 0, 1, ar, 0);
            writeBook(book2, thisWorkDirMag + "tfMagRep.xls");
            closeBook(book2);
        }

        openFile(thisWorkDirMag + "tfMagRep.xls");


    }

    //take unsorted tf chok report, make it sorted in order of passport
    public static void mainForTfReportChok() {

        String thisWorkDirShok = getWorkDir() + getTfReportsDir();
        String thisFileShok = thisWorkDirShok + "choc now.xls";

        HSSFWorkbook book = createBook(thisFileShok);

        ArrayList<HashMap> maplist = new ArrayList<HashMap>();

        for (int i = 0; i < 1; i++) {

            HashMap<String, Integer> map = createMapOrderSingleForReport(book, "sheet", 0, 0, 1);
            maplist.add(map);

        }


        HashMap<String, String> map2 = createMapOfArtsAndNamesBlankSpecial();

        HashMap<String, Integer> notContMap = new HashMap<String, Integer>();


        for (int e = 0; e < maplist.size(); e++) {

            HashMap<String, Integer> map = maplist.get(e);

            HashMap<String, Integer> newmap = new HashMap<String, Integer>();

            for (Map.Entry<String, Integer> entry : map.entrySet()) {

                String tfName = entry.getKey();
                Integer amount = entry.getValue();

                if (map2.containsKey(tfName)) {
                    newmap.put(map2.get(tfName), amount);
                } else {
                    notContMap.put(tfName, 0);
                }

            }


            ArrayList<ArrayList> list = createLongListOrderFromSingleMap(newmap, createListOfArtsFromPassport());

            printArray(list);

            writeVertListInBookSingle(book, String.valueOf(e), 0, 0, 1, list);

        }

        if (notContMap.size() > 0) {
            System.out.println("Нет в файле с наименованиями ТФ и нашими артикулами: ");
            printHashMap(notContMap);
            return;
        }
        writeBook(book, thisFileShok);
        closeBook(book);
        openFile(thisFileShok);

    }

    //take unsorted tf chok order, make it sorted in order of blank, write it on same file in sheet named "ФОРМА"
    public static void mainForTfOrderChok(String dir, String fileName) {

        String thisWorkDir = getWorkDir() + dir;

        //ArrayList data = createDataForMultipleOrder(dir);

        ArrayList data = new ArrayList();
        data.add("tf choc.xls");
        data.add("Sheet1");
        data.add("ФОРМА");
        data.add(1);
        data.add(1);
        data.add(2);
        data.add(1);


        String orderFileName = data.get(0).toString();
        String sheetName = data.get(1).toString();
        String outSheetName = data.get(2).toString();
        int firstRowNum = (Integer) data.get(3);
        int artColNum = (Integer) data.get(4);
        int firstAmColNum = (Integer) data.get(5);
        int step = (Integer) data.get(6);

        HSSFWorkbook book = createBook(thisWorkDir + orderFileName);

        HashMap<String, ArrayList> orderMap = createMapOrderMultipleSpecial(book, sheetName, firstRowNum, artColNum, firstAmColNum, step);

        HashMap<String, String> namesMap = createMapOfArtsAndNamesBlankSpecial();

        HashMap<String, ArrayList> newMap = new HashMap<String, ArrayList>();


        for (Map.Entry<String, ArrayList> entry : orderMap.entrySet()) {

            String goodsName = entry.getKey();

            if (namesMap.containsKey(goodsName)) {
                newMap.put(namesMap.get(goodsName), entry.getValue());
            }

        }

        closeBook(book);

        ArrayList list = createLongListOrderFromMultipleMap(newMap, createListOfArtsFromBlank());
        ArrayList remainsList = createShortListOrderFromMultipleMap(newMap, createListOfArtsFromPassport());


        HSSFWorkbook book2 = createBook(thisWorkDir + orderFileName);
        writeVertListInBookMultiple(book2, outSheetName, 1, 0, 1, list);
        writeRemainsListInBookMultiple(book2, outSheetName, remainsList);
        writeBook(book2, thisWorkDir + orderFileName);
        closeBook(book2);

    }

}
