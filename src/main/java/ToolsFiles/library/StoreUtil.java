package ToolsFiles.library;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.net.URI;
import java.nio.file.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreWriting.*;

public class StoreUtil {

    private static String passportName = "passport.xls";
    private static String clientsAndPricesName = "Клиенты и цены.xls";
    private static String blankName = "Бланк заказа Питер с допродажами.xls";
    //private static String blankName = "Бланк заказа Питер.xls";
    //private static String blankName = "Буквоед бланк.xls";

    private static String workDir = "C:/Java/Projects/JavaTest/";
    private static String clientsAndPricesDir = workDir;
    private static String blankDir = workDir;
    private static String passportDir = workDir;

    private static String clientsAndPricesFile = clientsAndPricesDir + clientsAndPricesName;
    private static String blankFile = blankDir + blankName;
    private static String passportFile = blankDir + passportName;

    private static String tfMagnetsCreateBigOrderDir = "tfMagnetsCreateBigOrder/";
    private static String updateFileForNewPassportDir = "updateFileForNewPassport/";
    private static String unsortToSortDir = "unsortToSort/";
    private static String singleOrderFromOldOrderBlankDir = "SingleOrderFromOldOrderBlank/";
    private static String singleOrderFromClientOrderDir = "SingleOrderFromClientOrder/";
    private static String multipleOrderFromClientOrderDir = "MultipleOrderFromClientOrder/";
    private static String tfReportsDir = "tfReports/";
    private static String commonSalesReportDir = "CommonSalesReport/";
    private static String remainsDir = "Remains/";
    private static String sortedOrderDir = "SortedOrder/";


    //----------------------------------PRINT CONSOLE

    //Printed all elements of ArrayList (toString)
    public static void printArray(ArrayList list) {


        for (int i = 0; i < list.size(); i++) {
            //System.out.println("i " + i + ", Class " + list.get(i).getClass() + ", value " + String.valueOf(list.get(i)));
            System.out.println("i " + i + ", value " + String.valueOf(list.get(i)));
        }
        System.out.println();
    }

    //Printed all elements of ArrayList of ArrayLists<String art, Int val> with names from certain Map (pasp or blank)
    public static void printStringArrayWhithNames(ArrayList list, HashMap<String, String> nameMap) {

        for (Object pos : list) {

            ArrayList p = (ArrayList) pos;
            String art = (String) p.get(0);

            if ((nameMap.containsKey(art)) && (p.size() == 2)) {
                String name = nameMap.get(art);
                if (p.get(1) == null) {
                    continue;
                }
                int am = (Integer) p.get(1);
                //System.out.println(art + " " + name + ", " + am + " шт");
            } else {
                System.out.println("there is no art " + art + " in list of articuls!");
            }

        }
    }

    //Printed all elements of HashMap
    public static <E, T> void printHashMap(HashMap<E, T> map) {

        //System.out.println("\nPrinting HashMap\n");

        for (Map.Entry entry : map.entrySet()) {
            if (entry.getKey() != null) {
                //System.out.println("Key " + entry.getKey() + ", value " + entry.getValue());
                System.out.println(entry.getKey().toString() + ", " + entry.getValue().toString());
            }
        }
        System.out.println();

    }

    //Printed all elements of HashMap<String, Integer> whith names from certain Map (pasp or blank)
    public static <E, T> void printHashMapWhithNames(HashMap<E, T> map, HashMap<String, String> nameMap) {

        for (Map.Entry entry : map.entrySet()) {
            String art = (String) entry.getKey();

            if ((entry.getKey() != null) && (nameMap.containsKey(entry.getKey()))) {
                String name = nameMap.get(entry.getKey());
                int am = (Integer) entry.getValue();

                System.out.println(art + " " + name + ", " + am + " шт");
            } else {
                System.out.println("there is no art " + art + " in map of articuls!");
            }
        }
    }


    //OTHER


    public static void copyFile(File fileFrom, File fileTo) {


        try {
            FileUtils.copyFile(fileFrom, fileTo);
        }
        catch (Exception e) {
            e.getMessage();
        }

    }


    //Open file
    public static void openFile(String fileName) {

        File file = new File(fileName);

        Desktop desktop = null;
        if (Desktop.isDesktopSupported()) {
            desktop = Desktop.getDesktop();
        }

        try {
            desktop.open(file);
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }

    }

    //Find and return our articul in String and return it
    public static String findArt(String rowArt) {

        String art = null;

        String pat1 = "([р|p|P|Р|Р][.](питер)(2)_(\\d{2}))";
        String pat2 = "([p|р|Р|P|Р](\\d+)_(\\d+))";
        String pat3 = "([p|р|P|Р|Р](\\d+)-(\\d+)-(\\d+))";
        String pat4 = "([р|p][.](москва)(2)_(\\d{2}))";
        String pat5 = "([р|p][.](выборг)(2)_(\\d{2}))";
        String pat6 = "([р|p][.](сочи)(2)_(\\d{2}))";
        String pat7 = "([C|С][-](22))";
        String pat8 = "([р|p|P|Р|Р][.](крым)(2)_(\\d{2}))";
        //String pat9 = "";
        String pattern = "(" + pat1 + "|" + pat2 + "|" + pat3 + "|" + pat4 + "|" + pat5 + "|" + pat6 + "|" + pat7 + "|" + pat8 + ")";

        Pattern p = Pattern.compile(pattern);
        Matcher m = p.matcher(rowArt);

        if (m.find()) {
            //System.out.println("Found " + m.group() + " in " + rowArt);
            art = m.group();

            if (art.contains("питер2_")) {

                int f = Integer.parseInt(art.substring(art.length() - 2, art.length()));

                if ((f > 26) && (f < 43)) {

                    art += "кр";
                    //System.out.println("After added :" + art);
                }
            }
        } else {
            System.out.println("Didn't find articul in " + rowArt);
        }

        return art;
    }

    // Take a art and map (from blank or passport) and return name of position
    public static String getNamePos(String art, HashMap<String, String> map) {

        if (map.containsKey(art)) {
            return map.get(art);
        } else {
            return null;
        }

    }

    //Print Array of String, return number of element that was choosen by user
    public static int chooseElementFromList(ArrayList<String> list) {
        for (int i = 0; i < list.size(); i++) {
            System.out.println(i + " " + list.get(i));
        }

        Scanner scanner = new Scanner(System.in);

        Integer n;

        do {

            try {
                System.out.println("Введите номер (Цифра от 0 до " + (list.size() - 1) + ")");
                n = Integer.parseInt(scanner.next());
            } catch (Exception e) {
                n = null;
            }
        }

        while (((n >= list.size()) || (n < 0) || !(n != null)));


        System.out.println("Выбран номер " + n + ", пункт " + list.get(n));

        return n;

    }

    // IN : ArrayList of values
    // DO : print to console ArrayList
    // and ask user a number
    // OUT: number of element of ArrayList
    public static int getNumberOfValueFromArray(ArrayList list) {
        for (int i = 0; i < list.size(); i++) {
            System.out.println(i + " " + list.get(i));
        }

        Scanner scanner = new Scanner(System.in);

        Integer n;

        do {

            try {
                System.out.println("Введите номер (Цифра от 0 до " + (list.size() - 1) + ")");
                n = Integer.parseInt(scanner.next());
            } catch (Exception e) {
                n = null;
            }
        }

        while (((n >= list.size()) || (n < 0) || (n == null)));


        System.out.println("Выбран пункт " + n + ", " + list.get(n));

        return n;

    }

    // Count summ of values from map
    public static <T> int mapValueSumm(HashMap<T, Integer> map) {
        int summ = 0;

        for (Map.Entry entry : map.entrySet()) {

            if (entry.getValue() == null) {
                continue;
            }
            summ += (Integer) entry.getValue();

        }

        return summ;
    }

    // Count summ of values from Array(String, Int)
    public static int arrValueSumm(ArrayList list) {
        int summ = 0;

        for (Object i : list) {

            ArrayList a = (ArrayList) i;
            int r = (Integer) a.get(1);
            summ += r;

        }

        return summ;
    }

    //Create list of files and directories from full directory name
    public static ArrayList<String> createDirectoryList(String directoryName) {
        //System.out.println(directoryName);
        ArrayList<String> list;
        File pathDir = new File(directoryName);
        String[] dirArray = pathDir.list();

        /*
        for (String s : dirArray)
        {list.add(s);}
        */

        list = new ArrayList<String>(Arrays.asList(dirArray));
        return list;

    }

    //Create list of sheet names from book
    public static ArrayList<String> createSheetsList(HSSFWorkbook book) {

        ArrayList<String> list = new ArrayList<String>();

        Iterator<Sheet> it = book.sheetIterator();

        while (it.hasNext()) {
            HSSFSheet s = (HSSFSheet) it.next();
            list.add(s.getSheetName());
        }
        return list;

    }

    //Create array of ALL files (reference FILE) from directory, including inner files
    public static ArrayList<File> createListOfAllFiles(ArrayList<File> ar, String dirString) {

        File dir = new File(dirString);

        ArrayList<File> arDir = new ArrayList<File>(Arrays.asList(dir.listFiles()));

        ListIterator<File> iter = arDir.listIterator();

        while (iter.hasNext()) {
            File f = iter.next();

            if (f.isDirectory()) {
                ArrayList<File> ar2 = createListOfAllFiles(ar, f.getAbsolutePath());
                iter.remove();
            }
        }

        ar.addAll(arDir);

        return ar;
    }

    //Modify list of book for actual passport (made for report file)
    public static void modifyListForActualPasport(HSSFWorkbook book, String sheetName, int firstRowNum, int artColNum, ArrayList<ArrayList> actualList) {

        HSSFSheet sheet = book.getSheet(sheetName);

        if (sheet == null) {
            System.out.println("Error: there is no sheet named " + sheetName + "!");
            return;
        }

        for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {

            HSSFRow row = sheet.getRow(i);

            if (row == null) {
                System.out.println("Error: row is null!");
                return;
            }

            HSSFCell cell = row.getCell(artColNum);

            if (cell == null) {
                //System.out.println("Error: cell is null!");
                return;
            }

            String art = cell.getStringCellValue();
            String artPasp = actualList.get(i - firstRowNum).get(0).toString();
            String namePasp = actualList.get(i - firstRowNum).get(1).toString();

            if (!art.equals(artPasp)) {
                System.out.println("Row " + i + " in sheet " + sheetName + " is not correspont for passport! value is " + art + ", replace by " + artPasp);
                sheet.shiftRows(i, sheet.getLastRowNum(), 1);
                row.createCell(artColNum).setCellValue(artPasp);
                row.createCell(artColNum + 1).setCellValue(namePasp);

            } else {
                continue;
            }

        }

    }

    // 0 - orderFileName
    // 1 - sheetname g
    // 2 - outSheetName
    // 3 - firstRowNum
    // 4 - artColNum
    // 5 - firstAmColNum
    // 6 - step
    //
    public static ArrayList createDataForSorting(String workingDirectory) {

        ArrayList dataArray = new ArrayList(8);
        Scanner s = new Scanner(System.in);

        // Get a order file name
        ArrayList<String> dirAr = createDirectoryList(getWorkDir() + workingDirectory);

        System.out.println("Выбор файла с заказом");
        String orderFileName = dirAr.get(chooseElementFromList(dirAr));
        dataArray.add(orderFileName);


        HSSFWorkbook orderBook = createBook(getWorkDir() + workingDirectory + orderFileName);

        // Get a sheet name
        System.out.println("\nВыбор листа, на котором содержится заказ");
        ArrayList<String> sheetsAr = createSheetsList(orderBook);
        String sheetName = sheetsAr.get(chooseElementFromList(sheetsAr));
        dataArray.add(sheetName);

        closeBook(orderBook);

        String outSheetName = sheetName + " sorted";
        dataArray.add(outSheetName);

        //get a row num
        System.out.println("\nВведите номер строки, с которой начинается таблица заказа");
        int firstRowNum = s.nextInt();
        dataArray.add(firstRowNum);

        //get a art col
        System.out.println("\nВведите номер колонки с данными, содержащими артикул");
        int artColNum = s.nextInt();
        dataArray.add(artColNum);

        //get a first am col
        System.out.println("\nВведите номер первой колонки с цифрами заказа");
        int firstAmColNum = s.nextInt();
        dataArray.add(firstAmColNum);

        //get a step
        System.out.println("\nВведите шаг заказа");
        int step = s.nextInt();
        dataArray.add(step);

        //get a mask (1 - mask is blank, 2 - mask is passport)
        System.out.println("\nВведите 1, если шаблон - бланк заказа, или 2, если шаблон - паспорт");
        int mask = s.nextInt();
        dataArray.add(mask);

        System.out.println("\nДанные приняты. Отсортированный заказ будет записан в тот же файл," +
                "на лист \"" + outSheetName + "\"");

        return dataArray;
    }

    // 0 - orderFileName
    // 1 - sheetname g
    // 2 - outSheetName
    // 3 - firstRowNum
    // 4 - artColNum
    // 5 - firstAmColNum
    // 6 - step
    //
    public static ArrayList createDataForNewOrder(String workingDirectory) {

        ArrayList dataArray = new ArrayList(8);
        Scanner s = new Scanner(System.in);

        // Get a order file name
        ArrayList<String> dirAr = createDirectoryList(getWorkDir() + workingDirectory);

        System.out.println("Выбор файла с заказом");
        String orderFileName = dirAr.get(chooseElementFromList(dirAr));
        dataArray.add(orderFileName);


        HSSFWorkbook orderBook = createBook(getWorkDir() + workingDirectory + orderFileName);

        // Get a sheet name
        System.out.println("\nВыбор листа, на котором содержится заказ");
        ArrayList<String> sheetsAr = createSheetsList(orderBook);
        String sheetName = sheetsAr.get(chooseElementFromList(sheetsAr));
        dataArray.add(sheetName);

        closeBook(orderBook);

        String outSheetName = sheetName;
        dataArray.add(outSheetName);

        //get a row num
        dataArray.add(9);

        //get a art col
        dataArray.add(0);

        //get a first am col
        dataArray.add(4);

        //get a step
        dataArray.add(1);

        //get a mask (1 - mask is blank, 2 - mask is passport)
        dataArray.add(1);


        return dataArray;
    }

    // 0 - orderFileName
    // 1 - sheetname g
    // 2 - outSheetName
    // 3 - firstRowNum
    // 4 - artColNum
    // 5 - firstAmColNum
    // 6 - step
    //
    public static ArrayList createDataForOrderFromClientOrder(String workingDirectory) {

        ArrayList dataArray = new ArrayList(6);
        Scanner s = new Scanner(System.in);

        // Get a order file name
        ArrayList<String> dirAr = createDirectoryList(getWorkDir() + workingDirectory);

        System.out.println("Выбор файла с заказом");
        String orderFileName = dirAr.get(chooseElementFromList(dirAr));
        dataArray.add(orderFileName);


        HSSFWorkbook orderBook = createBook(getWorkDir() + workingDirectory + orderFileName);

        // Get a sheet name
        System.out.println("\nВыбор листа, на котором содержится заказ");
        ArrayList<String> sheetsAr = createSheetsList(orderBook);
        String sheetName = sheetsAr.get(chooseElementFromList(sheetsAr));
        dataArray.add(sheetName);

        closeBook(orderBook);

        dataArray.add("ФОРМА");

        //get a row num
        System.out.println("\nВведите номер строки, с которой начинается таблица заказа");
        int firstRowNum = s.nextInt();
        dataArray.add(firstRowNum);

        //get a art col
        System.out.println("\nВведите номер колонки с данными, содержащими артикул");
        int artColNum = s.nextInt();
        dataArray.add(artColNum);

        //get a first am col
        System.out.println("\nВведите номер первой колонки с цифрами заказа");
        int firstAmColNum = s.nextInt();
        dataArray.add(firstAmColNum);


        return dataArray;
    }

    // 0 - orderFileName
    // 1 - sheetname g
    // 2 - outSheetName
    // 3 - firstRowNum
    // 4 - artColNum
    // 5 - firstAmColNum
    // 6 - step
    //
    public static ArrayList createDataForMultipleOrder(String workingDirectory) {

        ArrayList dataArray = new ArrayList(8);
        Scanner s = new Scanner(System.in);

        // Get a order file name
        ArrayList<String> dirAr = createDirectoryList(getWorkDir() + workingDirectory);

        System.out.println("Выбор файла с заказом");
        String orderFileName = dirAr.get(chooseElementFromList(dirAr));
        dataArray.add(orderFileName);


        HSSFWorkbook orderBook = createBook(getWorkDir() + workingDirectory + orderFileName);

        // Get a sheet name
        System.out.println("\nВыбор листа, на котором содержится заказ");
        ArrayList<String> sheetsAr = createSheetsList(orderBook);
        String sheetName = sheetsAr.get(chooseElementFromList(sheetsAr));
        dataArray.add(sheetName);

        closeBook(orderBook);

        String outSheetName = "ФОРМА";
        dataArray.add(outSheetName);

        //get a row num
        System.out.println("\nВведите номер строки, с которой начинается таблица заказа");
        int firstRowNum = s.nextInt();
        dataArray.add(firstRowNum);

        //get a art col
        System.out.println("\nВведите номер колонки с данными, содержащими артикул");
        int artColNum = s.nextInt();
        dataArray.add(artColNum);

        //get a first am col
        System.out.println("\nВведите номер первой колонки с цифрами заказа");
        int firstAmColNum = s.nextInt();
        dataArray.add(firstAmColNum);

        //get a step
        System.out.println("\nВведите шаг заказа");
        int step = s.nextInt();
        dataArray.add(step);

        //get a mask (1 - mask is blank, 2 - mask is passport)
        dataArray.add(1);

        return dataArray;
    }

    //////////////////////////////////////// CREATING, WRITING, CLOSING BOOKS

    public static HSSFWorkbook createBook(String filename) {

        try {
            HSSFWorkbook book = new HSSFWorkbook(new FileInputStream(filename));
            return book;
        } catch (IOException e) {
            System.out.println("Book " + filename + " was not created!\n");
            e.printStackTrace();
        }

        return null;

    }

    public static XSSFWorkbook createXBook(String filename) {

        try {
            XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(filename));
            //System.out.println("Book " + filename + " was created\n");
            return book;
        } catch (IOException e) {
            System.out.println("Book was not created!\n");
            e.printStackTrace();
        }

        return null;

    }

    public static void writeBook(HSSFWorkbook book, String bookName) {
        try {
            FileOutputStream fos = new FileOutputStream(bookName);
            book.write(fos);
            fos.close();
            //System.out.println("Book " + bookName + " is wroted!\n");
        } catch (IOException e) {
            System.out.println("Warning! Book " + bookName + " is NOT wroted!\n");
            closeBook(book);
            e.printStackTrace();
        }

    }

    public static void writeXBook(XSSFWorkbook book, String bookName) {
        try {
            FileOutputStream fos = new FileOutputStream(bookName);
            book.write(fos);
            fos.close();
            //  System.out.println("Book " + bookName + " is wroted!\n");
        } catch (IOException e) {
            System.out.println("Warning! Book " + bookName + " is NOT wroted!\n");
            e.printStackTrace();
        }

    }

    public static void closeBook(HSSFWorkbook book) {

        try {
            book.close();
            //  System.out.println("Book was closed\n");
        } catch (IOException e) {
            System.out.println("Book is not closed!\n");
            e.printStackTrace();
        }
    }

    public static void closeXBook(XSSFWorkbook book) {

        try {
            book.close();
            //  System.out.println("Book was closed\n");
        } catch (IOException e) {
            System.out.println("Book is not closed!\n");
            e.printStackTrace();
        }
    }


    /////////////////////////////////////// GETTERS


    public static String getClientsAndPricesName() {
        return clientsAndPricesName;
    }

    public static String getBlankName() {
        return blankName;
    }

    public static String getClientsAndPricesFile() {
        return clientsAndPricesFile;
    }

    public static String getBlankFile() {
        return blankFile;
    }

    public static String getPassportFile() {
        return passportFile;
    }


    public static String getWorkDir() {
        return workDir;
    }

    public static String getClientsAndPricesDir() {
        return clientsAndPricesDir;
    }

    public static String getBlankDir() {
        return blankDir;
    }

    public static String getTfMagnetsCreateBigOrderDir() {
        return tfMagnetsCreateBigOrderDir;
    }

    public static String getTfReportsDir() {
        return tfReportsDir;
    }

    public static String getUpdateFileForNewPassportDir() {
        return updateFileForNewPassportDir;
    }

    public static String getUnsortToSortDir() {
        return unsortToSortDir;
    }

    public static String getSingleOrderFromOldOrderBlankDir() {
        return singleOrderFromOldOrderBlankDir;
    }

    public static String getSingleOrderFromClientOrderDir() {
        return singleOrderFromClientOrderDir;
    }

    public static String getMultipleOrderFromClientOrderDir() {
        return multipleOrderFromClientOrderDir;
    }

    public static String getCommonSalesReportDir() {
        return commonSalesReportDir;
    }

    public static String getRemainsDir() {
        return remainsDir;
    }

    public static String getPassportDir() {
        return passportDir;
    }

    public static String getSortedOrderDir() {
        return sortedOrderDir;
    }


///////////////////////   GARBAGE

    public static void saveNewBlank() {

        try {
            FileWriter fw = new FileWriter(getBlankDir() + blankName);
            //fw.append(blankDir + blankName);
            fw.close();
            System.out.println("Файл создан!");
        } catch (IOException e) {
            System.out.println("error!!");
            e.printStackTrace();
        }

    }

    public static ArrayList getPrices(HSSFWorkbook book, int columnNum) {

        ArrayList list = new ArrayList();
        HSSFSheet sheet = book.getSheet("Цены");

        for (int i = 3; i < 5; i++) {

            HSSFRow row = sheet.getRow(i);
            HSSFCell cell = row.getCell(columnNum);
            String s = cell.getStringCellValue();
            list.add(s);
        }

        for (int i = 5; i < 33; i++) {

            HSSFRow row = sheet.getRow(i);
            HSSFCell cell = row.getCell(columnNum);

            double s = cell.getNumericCellValue();

            double r = s - (int) s;

            if (r == 0) {

                //System.out.println("s " + s + ", (int)s " + (int)s + ", r = " + (s - (int)s) );
                int ns = (int) s;
                //System.out.println("ns = " + (int)s);
                list.add(ns);
            } else {
                list.add(s);
            }
        }

        HSSFRow row = sheet.getRow(33);
        HSSFCell cell = row.getCell(columnNum);

        list.add((long) cell.getNumericCellValue());

        printArray(list);

        return list;
    }

    public static void switching(int i) {

        switch (i) {


            case 0:
                System.out.println("EXIT");
                break;
            case 1:
                //TFMagnetsCreateBigOrder.main("tf order.xls", "order", "tf order.xls", "TFCreateOrders/");
                break;
            case 2:
                //OldToNewOrder.main();
                break;
            case 3:
                //MakeOrderFromUnsortedOrder.main("ДК заказ.xls", "sortedorder", "ДК заказ.xls", "sheet", 0);
                break;
            case 4:
                //MakeClientOrderBlank.main();
                break;
            case 5:
                ShipmentStore.main("Report/", "commonReport.xls", 2);
                break;
            default:
                System.out.println("Неправильное значение!");
                break;
        }

    }

    public static void cellStyleTest() {

        String bookName = getWorkDir() + "styleTest.xls";
        String sheetName = "sheet";

        HSSFWorkbook book = createBook(bookName);
        HSSFSheet sheet = book.getSheet(sheetName);
        HSSFRow row = sheet.getRow(0);
        HSSFCell cell = row.getCell(0);


        //cell.setCellStyle(style);

        writeBook(book, bookName);
        closeBook(book);

    }

    public static int choose() {

        System.out.println("Что делаем?");
        System.out.println("1 - Сортируем заказ ТФ магниты");
        System.out.println("2 - Наш заказ из чужого (старый бланк)");
        System.out.println("3 - Наш заказ из чужого несортированного");
        System.out.println("4 - Бланк заказа клиента");
        System.out.println("5 - Отчет по отгрузкам");
        System.out.println("0 - Выйти из программы");

        Scanner s = new Scanner(System.in);
        return s.nextInt();

    }

    public static void dirTest() {

        ArrayList<File> ar = createListOfAllFiles(new ArrayList<File>(), getWorkDir());


        for (File f : ar) {
            System.out.println(f.getName());
        }

        System.out.println(ar.size());

    }

    public static void mapListTest() {

        String sheetSingleOrderName = "testS";
        String sheetMultipleOrderName = "testM";

        HSSFWorkbook book = createBook(getWorkDir() + "GermesOrder.xls");
        HashMap<String, Integer> singleOrderMap1 = createMapOrderSingle(book, sheetSingleOrderName, 0, 0, 1);
        HashMap<String, Integer> singleOrderMap2 = createMapOrderSingle(book, sheetSingleOrderName, 0, 0, 1);
        HashMap<String, Integer> singleOrderMap3 = createMapOrderSingle(book, sheetSingleOrderName, 0, 0, 1);
        HashMap<String, Integer> singleOrderMap4 = createMapOrderSingle(book, sheetSingleOrderName, 0, 0, 1);
        HashMap<String, ArrayList> multipleOrderMap1 = createMapOrderMultiple(book, sheetMultipleOrderName, 0, 0, 1, 1);
        HashMap<String, ArrayList> multipleOrderMap2 = createMapOrderMultiple(book, sheetMultipleOrderName, 0, 0, 1, 1);
        HashMap<String, ArrayList> multipleOrderMap3 = createMapOrderMultiple(book, sheetMultipleOrderName, 0, 0, 1, 1);
        HashMap<String, ArrayList> multipleOrderMap4 = createMapOrderMultiple(book, sheetMultipleOrderName, 0, 0, 1, 1);
        closeBook(book);

        //System.out.println("Single Map");
        //printHashMap(singleOrderMap1);
        //System.out.println("Multiple Map");
        //printHashMap(multipleOrderMap1);

        ArrayList<String> singleListShort = createVertListOfValues(book, "testS", 0, 0);
        ArrayList<ArrayList> singleOrderListShort = createShortListOrderFromSingleMap(singleOrderMap1, createListOfArtsFromBlank());
        ArrayList<ArrayList> singleOrderListLong = createLongListOrderFromSingleMap(singleOrderMap2, createListOfArtsFromBlank());
        ArrayList<ArrayList> orderListShortWithNames = createShortListOrderFromSingleMapWithNames(singleOrderMap3, createListOfArtsAndNamesBlank());
        ArrayList<ArrayList> orderListLongWithNames = createLongListOrderFromSingleMapWithNames(singleOrderMap4, createListOfArtsAndNamesBlank());
        ArrayList<ArrayList> multipleOrderListShort = createShortListOrderFromMultipleMap(multipleOrderMap1, createListOfArtsFromBlank());
        ArrayList<ArrayList> multipleOrderListLong = createLongListOrderFromMultipleMap(multipleOrderMap2, createListOfArtsFromBlank());
        ArrayList<ArrayList> multipleOrderListShortWithNames = createShortListOrderFromMultipleMapWithNames(multipleOrderMap3, createListOfArtsAndNamesBlank());
        ArrayList<ArrayList> multipleOrderListLongWithNames = createLongListOrderFromMultipleMapWithNames(multipleOrderMap4, createListOfArtsAndNamesPassport());

                        /*
        System.out.println("Single Short List");
        printArray(singleOrderListShort);

        System.out.println("Single Long List");
        printArray(singleOrderListLong);
        System.out.println("Single Short List with Names");
        printArray(orderListShortWithNames);
        System.out.println("Single Long List with Names");
        printArray(orderListLongWithNames);
        System.out.println("Multiple Short List");
        printArray(multipleOrderListShort);
        System.out.println("Multiple Long List");
        printArray(multipleOrderListLong);
        System.out.println("Multiple Short List with Names");
        printArray(multipleOrderListShortWithNames);
        System.out.println("Multiple Long List with Names");
        printArray(multipleOrderListLongWithNames);

        */

        HSSFWorkbook book2 = createBook(getWorkDir() + "GermesOrder.xls");

        /*
        writeArrayList(book2, "singleListShort", 0, 0, singleListShort);
        writeStringArrayList(book2, "singleListShort2", 0, 0, singleListShort);

        writeArrayArrayStringInt(book2, "singleOrderListShort", 0, 0,1, singleOrderListShort, "");
        writeArrayArrayStringInt(book2, "singleOrderListShort0", 0, 0,1,singleOrderListShort, 0);
        writeArrayArrayStringInt(book2, "singleOrderListShortNull", 0, 0,1, singleOrderListShort, null);

        writeArrayArrayStringInt(book2, "singleOrderListLong", 0, 0,1, singleOrderListLong, "");
        writeArrayArrayStringInt(book2, "singleOrderListLong0", 0, 0,1,singleOrderListLong, 0);
        writeArrayArrayStringInt(book2, "singleOrderListLongNull", 0, 0,1, singleOrderListLong, null);


        writeArrayArrayStringNameInt(book2, "orderListLongWithNames", 0, 0,1,2, orderListLongWithNames, "");
        writeArrayArrayStringNameInt(book2, "orderListLongWithNames0", 0, 0,1,2,orderListLongWithNames, 0);
        writeArrayArrayStringNameInt(book2, "orderListLongWithNamesNull", 0, 0,1,2, orderListLongWithNames, null);


        writeArrayArrayStringArray(book2, "multipleOrderListLong", 0, 0,1, multipleOrderListLong, "");
        writeArrayArrayStringArray(book2, "multipleOrderListLong0", 0, 0,1,multipleOrderListLong, 0);
        writeArrayArrayStringArray(book2, "multipleOrderListLongNull", 0, 0,1,multipleOrderListLong, null);
*/

        writeArrayArrayStringNameArray(book2, "multipleOrderListLongWithNames", 0, 0, 1, 2, multipleOrderListLongWithNames, "");
        writeArrayArrayStringNameArray(book2, "multipleOrderListLongWithNames0", 0, 0, 1, 2, multipleOrderListLongWithNames, 0);
        writeArrayArrayStringNameArray(book2, "multipleOrderListLongWithNamesNull", 0, 0, 1, 2, multipleOrderListLongWithNames, null);


        writeBook(book2, getWorkDir() + "GermesOrder.xls");
        closeBook(book2);
        openFile(getWorkDir() + "GermesOrder.xls");

    }

    public static int getClient(ArrayList<String> list) {
        for (int i = 0; i < list.size(); i++) {
            System.out.println(i + " " + list.get(i));
        }

        Scanner scanner = new Scanner(System.in);

        Integer n;

        do {

            try {
                System.out.println("Введите номер клиента (Цифра от 0 до " + (list.size() - 1) + ")");
                n = Integer.parseInt(scanner.next());
            } catch (Exception e) {
                n = null;
            }
        }

        while (((n >= list.size()) || (n < 0) || !(n != null)));


        System.out.println("Клиент " + n + ", " + list.get(n));

        return n;

    }

}
