package ToolsFiles.library;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;

import static ToolsFiles.library.StoreUtil.*;
import static ToolsFiles.library.StoreCreating.*;
import static ToolsFiles.library.StoreWriting.*;
import static ToolsFiles.library.Client.*;

public class ClientParse {

    static String name, contact;
    static long inn;

    public static void main(String[] args) {

        ArrayList<Client> clAr = new ArrayList<Client>();
        HSSFWorkbook book = createBook(getClientsAndPricesFile());

        HSSFSheet sheet = book.getSheet("Данные");

        for (int i = 1; i < 10; i++) {
            name = sheet.getRow(i).getCell(6).getStringCellValue();
            inn = (long) sheet.getRow(i).getCell(8).getNumericCellValue();
            contact = "";

            if (sheet.getRow(i + 1).getCell(6).getStringCellValue() == "") {
                System.out.println(i + ", name is null");

                do {
                    String contactN = sheet.getRow(i).getCell(10).getStringCellValue();
                    String contactF = sheet.getRow(i).getCell(11).getStringCellValue();
                    if(contactF == "") {
                        System.out.println(i + ", contact no number");
                        contactF = "no phone number";}
                    if (contact.length()>0) {contact = contact + ",\n" + contactN + ", " + contactF;}
                    else {contact = contactN + ", " + contactF;}
                    i++;
                }
                while (sheet.getRow(i + 1).getCell(6).getStringCellValue() == "");

            }

            clAr.add(new Client(name, inn, contact));
        }

        closeBook(book);

        for (Client c : clAr) {
            System.out.println(c);
        }
    }


}
