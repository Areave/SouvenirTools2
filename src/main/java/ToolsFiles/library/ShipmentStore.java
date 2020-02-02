package ToolsFiles.library;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import static ToolsFiles.library.StoreUtil.*;
import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreWriting.*;

public class ShipmentStore {

    // Write shipment in book
    // Way 1 - based on INN
    // Way 2 - common report
    // Way 3 - based on Cash or Transfer

    public static void main(String directoryName, String reportName, int way) {

        String reportFile = getWorkDir() + reportName;

        ArrayList<Shipment> shipList = createListShipmentFromDirectory(directoryName);

        HSSFWorkbook book = createBook(reportFile);

        for (Shipment ship : shipList) {
            writeShipmentInBook(book, ship, way);
        }


        writeBook(book, reportFile);
        closeBook(book);


    }

    // Create Shipment object from file
    public static Shipment createShipment(HSSFWorkbook book) {

        HSSFSheet sheet = book.getSheetAt(0);

        String a = sheet.getRow(2).getCell(1).getStringCellValue();
        if (sheet.getRow(2) == null) {
            System.out.println("Customer name is null");
        }

        if (sheet.getRow(2).getCell(1) == null) {
            System.out.println("cell null!");
        }


        String b = "";
        try {
            b = sheet.getRow(4).getCell(1).getStringCellValue();
        } catch (Exception e) {
        }

        String customer = a + " " + b;


        Date shipDate = sheet.getRow(5).getCell(1).getDateCellValue();

        if (shipDate == null) {
            System.out.println("Date is null!");

        }


        long inn = (long) sheet.getRow(3).getCell(1).getNumericCellValue();

        boolean isCash = false;
        if (inn == 7706092528l) {
            isCash = true;
        }

        double summ = 0;

        for (int i = sheet.getLastRowNum(); i > 0; i--) {
            if (sheet.getRow(i) == null) {continue;}
            HSSFCell cell = sheet.getRow(i).getCell(5);
            if (cell == null) {
                //System.out.println("i " + i + ", cont!");
                continue;
            } else {
                //cell.setCellType(CellType.NUMERIC);
                summ = cell.getNumericCellValue();
                break;
            }


        }

        //ArrayList<ArrayList> order = createListOrderShort(book, sheet.getSheetName(), 9, 0, 4);
        HashMap<String, Integer> order = createMapOrderSingle(book, sheet.getSheetName(), 9, 0, 4);


        if (summ == 0.0) {
            System.out.println("Error! " + customer + ", summ = 0!!!!");
        }

        return new Shipment(customer, shipDate, order, inn, isCash, summ);


    }

    // Create Shipment object from file for shortShipment
    public static Shipment createShipmentForCashShipmentShort(HSSFWorkbook book) {

        HSSFSheet sheet = book.getSheetAt(0);

        String a = sheet.getRow(2).getCell(1).getStringCellValue();
        if (sheet.getRow(2) == null) {
            System.out.println("Customer name is null");
        }

        if (sheet.getRow(2).getCell(1) == null) {
            System.out.println("cell null!");
        }


        String customer = a;


        Date shipDate = sheet.getRow(5).getCell(1).getDateCellValue();

        if (shipDate == null) {
            System.out.println("Date is null!");

        }


        long inn = (long) sheet.getRow(3).getCell(1).getNumericCellValue();

        boolean isCash = false;
        if (inn == 7706092528l) {
            isCash = true;
        }

        double summ = 0;

        for (int i = sheet.getLastRowNum(); i > 0; i--) {
            HSSFCell cell = sheet.getRow(i).getCell(5);
            if ((cell == null) || (sheet.getRow(i) == null)) {
                //System.out.println("i " + i + ", cont!");
                continue;
            } else {
                cell.setCellType(CellType.NUMERIC);
                summ = cell.getNumericCellValue();
                break;
            }


        }

        //ArrayList<ArrayList> order = createListOrderShort(book, sheet.getSheetName(), 9, 0, 4);
        HashMap<String, Integer> order = createMapOrderSingle(book, sheet.getSheetName(), 9, 0, 4);


        if (summ == 0.0) {
            System.out.println("Error! " + customer + ", summ = 0!!!!");
        }

        return new Shipment(customer, shipDate, order, inn, isCash, summ);


    }


    // Write shipment in book
    // Way 1 - based on INN
    // Way 2 - common report
    // Way 3 - based on Cash or Transfer
    public static void writeShipmentInBook(HSSFWorkbook book, Shipment ship, int way) {

        String sheetName = "";

        switch (way) {
            case 1:
                if (ship.getInn() == 7706092528l) {
                    sheetName = "Cash";
                } else {
                    sheetName = findNameByInn(ship.getInn(), createMapInn());

                    if (sheetName.length() > 30) {
                        sheetName = sheetName.substring(0, 20);
                    }
                }
                break;
            case 2:
                sheetName = "CommonReport";
                break;
            case 3:
                if (ship.getInn() == 7706092528l) {
                    sheetName = "Cash";
                } else {
                    sheetName = "BankTransfer";
                }
                break;

            case 4:
                if (ship.getInn() == 782600082016l) {
                    sheetName = "Airport";

                }
                else{return;}

                break;

            default:
                System.out.println("Way must be 1, 2 or 3 !");
                return;

        }

        writeVertListInBookSingleForReport(book, sheetName, createLongListOrderFromMultipleMap(ship.getOrder(), createListOfArtsFromPassport()), ship);
    }

    // Print shipmet in console
    public static void printShipment(Shipment ship) {


        //System.out.println("Customer: " + ship.getCustomer() + ", Date: " + ship.getData() + ", Summ: " + ship.getSumm());
        // System.out.println("Summ: " + ship.getSumm() + ", INN: " + ship.getInn() + ", cash: " + ship.getIsCash());

        // Deprecated method!
        //HashMap<String, Integer> newOrder = makeRemainsMapSinge(ship.getOrder(), createMapOfArtsAndNamesPassport());
        System.out.println();

        SimpleDateFormat format = new SimpleDateFormat("dd MMMM");
        String data = format.format(ship.getShipDate());

        System.out.println(ship.getCustomer() + ", " + data + ", " + ship.getSumm() + " рублей");
        ArrayList<ArrayList> orderList = createLongListOrderFromMultipleMap(ship.getOrder(), createListOfArtsFromPassport());
        printStringArrayWhithNames(orderList, createMapOfArtsAndNamesPassport());
        System.out.println();

    }

    // Pick random shipment from list of shipments
    public static Shipment takeRandomShipment(ArrayList<Shipment> shipList) {

        int r = new Random().nextInt(shipList.size());
        return shipList.get(r);

    }

    // Make HashMap (Inn, Name Of Client> from ClientAndPrices
    // Inn for cash is one for all
    public static HashMap<Long, String> createMapInn() {

        HSSFWorkbook book = createBook(getClientsAndPricesFile());
        HSSFSheet sheet = book.getSheet("Цены");
        HashMap<Long, String> map = new HashMap<Long, String>();


        for (int i = 3; i < sheet.getRow(3).getLastCellNum(); i++) {

            HSSFRow rowCl = sheet.getRow(4);
            if (rowCl.getCell(i) == null) {
                //System.out.println("NULL! " + i);
                break;
            }
            String client = rowCl.getCell(i).getStringCellValue();
            HSSFRow rowInn = sheet.getRow(33);
            Long inn = (long) rowInn.getCell(i).getNumericCellValue();

            //System.out.println("Put " + inn + ", " + client);
            map.put(inn, client);
        }

        closeBook(book);

        return map;
    }

    public static ArrayList<Shipment> createListShipmentFromDirectory(String workDirectory) {

        String workDirShip = getWorkDir() + workDirectory;

        ArrayList<String> dirList = createDirectoryList(workDirShip);

        ArrayList<Shipment> shipments = new ArrayList<Shipment>();

        for (String fileName : dirList) {
            // System.out.println(fileName);
            HSSFWorkbook book = createBook(workDirShip + fileName);
            System.out.println("Writing book " + fileName);
            Shipment ship = createShipment(book);
            closeBook(book);
            shipments.add(ship);
        }

        Collections.sort(shipments);

        return shipments;

    }

    public static String findNameByInn(long inn, HashMap<Long, String> mapInn) {

        if (mapInn.containsKey(inn)) {
            return mapInn.get(inn);
        } else {
            return String.valueOf(inn);
        }


    }

    public static void mergeFilesToAllDirectory(String commonDirName) {

        String dirCash = "cash/";
        String dirBank = "bank/";
        File allDir = new File(commonDirName + "all/");

        ArrayList<File> arCash = createListOfAllFiles(new ArrayList<File>(), commonDirName + dirCash);
        ArrayList<File> arBank = createListOfAllFiles(new ArrayList<File>(), commonDirName + dirBank);

        ArrayList<File> arAll = new ArrayList<File>();
        arAll.addAll(arBank);
        arAll.addAll(arCash);

        System.out.println(arAll.size());

        //Collections.sort(arAll);

        ListIterator<File> lit = arAll.listIterator();

        while (lit.hasNext()) {

            File f = lit.next();

            File nf = new File(allDir + "/" + f.getName());

            if (nf.exists()) {
                System.out.println("Exist! " + nf.getName());
                File nf2 = new File(allDir + "/" + returnUnikFileName(nf.getName(), new File(allDir + "/")));
                System.out.println("nf2.getName() " + nf2.getName());
                nf.renameTo(nf2);
                System.out.println("nf.getName() " + nf.getName());
            }

            try {
                FileUtils.copyFileToDirectory(f, allDir);
            } catch (IOException t) {
                t.printStackTrace();
            }

        }
    }

    public static void pullFilesToOneDirectory(String dir, String destDir) {

        File allDir = new File(destDir);

        ArrayList<File> arOfFiles = createListOfAllFiles(new ArrayList<File>(), dir);

        System.out.println("Files: " + arOfFiles.size());

        ListIterator<File> lit = arOfFiles.listIterator();

        while (lit.hasNext()) {

            File f = lit.next();

            File nf = new File(allDir + "/" + f.getName());

            if (nf.exists()) {
                System.out.println("Exist! " + nf.getName());
                File nf2 = new File(allDir + "/" + returnUnikFileName(nf.getName(), new File(allDir + "/")));
                System.out.println("nf2.getName() " + nf2.getName());
                nf.renameTo(nf2);
                System.out.println("nf.getName() " + nf.getName());
            }

            try {
                FileUtils.copyFileToDirectory(f, allDir);
            } catch (IOException t) {
                t.printStackTrace();
            }

        }
    }

    public static String returnUnikFileName(String fileName, File dir) {

        File newFile = new File(dir + "/" + fileName);
        System.out.println("newFile.getAbsolutePath() " + newFile.getAbsolutePath());
        if (newFile.exists()) {

            System.out.println("Exist inside method! " + newFile.getName());
            return returnUnikFileName("old_" + fileName, dir);

        }

        System.out.println("return " + fileName);
        return fileName;
    }
}


