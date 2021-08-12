package pits;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Locale;
import java.util.Scanner;

public class OrderProcessingMain {

    public static void main(String[] args) {

        try {

            Scanner scanner = new Scanner(System.in);
            boolean isEnviornmentUAT = true;

            System.out.println("Select the environment : (1/2)");
            System.out.println("1. UAT ");
            System.out.println("2. PROD");

            if (scanner.nextLine().equalsIgnoreCase("2")) {
                isEnviornmentUAT = false;
            }

            XSSFWorkbook orderWorkBook = importOrderWorkBook(scanner);

            XSSFWorkbook orderEntryWorkBook = importOrderEntryWorkBook(scanner);

            XSSFWorkbook addressWorkBook = importAddressWorkBook(scanner);

            createOrderAndAddressImpexFile(orderWorkBook, addressWorkBook, isEnviornmentUAT);

            createOrderEntryImpexFile(orderEntryWorkBook);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static XSSFWorkbook importOrderWorkBook(Scanner scanner) throws IOException {

        System.out.println("Enter Order Workbook name with extension : ");
        String fileName = scanner.nextLine();

        FileInputStream fileInputStream = new FileInputStream("./Source Folder/" + fileName);
        // customerFileInputStream.close();

        return new XSSFWorkbook(fileInputStream);
    }

    private static XSSFWorkbook importOrderEntryWorkBook(Scanner scanner) throws IOException {

        System.out.println("Enter Order Entry Workbook name with extension : ");
        String fileName = scanner.nextLine();

        FileInputStream fileInputStream = new FileInputStream("./Source Folder/" + fileName);
        // customerFileInputStream.close();

        return new XSSFWorkbook(fileInputStream);
    }

    private static XSSFWorkbook importAddressWorkBook(Scanner scanner) throws IOException {

        System.out.println("Enter Address Workbook name with extension : ");
        String fileName = scanner.nextLine();

        FileInputStream fileInputStream = new FileInputStream("./Source Folder/" + fileName);
        // customerFileInputStream.close();

        return new XSSFWorkbook(fileInputStream);
    }

    private static void createOrderAndAddressImpexFile(
            XSSFWorkbook orderWorkbook, XSSFWorkbook addressWorkBook, boolean isEnviornmentUAT) {

        CSVPrinter csvPrinter = null;
        try {

            csvPrinter =
                    new CSVPrinter(
                            new FileWriter("./Target Folder/OrderImpex.impex"),
                            CSVFormat.EXCEL.withDelimiter(';').withTrim().withQuoteMode(QuoteMode.MINIMAL));

            exportOrderDataToImpex(csvPrinter, orderWorkbook, isEnviornmentUAT);

            csvPrinter.println();

            exportAddressDataToImpex(csvPrinter, addressWorkBook);

        } catch (Exception e) {
            System.out.println("Failed to write Order and Address Impex file to output stream : ");
            e.printStackTrace();
        } finally {
            try {
                if (csvPrinter != null) {
                    csvPrinter.flush(); // Flush and close CSVPrinter
                    csvPrinter.close();
                }
            } catch (IOException ioe) {
                System.out.println("Error when closing CSV Printer");
            }
        }
    }

    private static void exportOrderDataToImpex(
            CSVPrinter csvPrinter, XSSFWorkbook orderWorkbook, boolean isEnviornmentUAT)
            throws IOException {
        if (orderWorkbook != null) {
            XSSFSheet orderSheet = orderWorkbook.getSheet("Order");

            Row headerRow = orderSheet.getRow(2);
            Iterator<Cell> cellIterator = headerRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (null != cell && cell.getCellType() != CellType.BLANK) {
                    csvPrinter.print(cell.toString());
                }
            }
            //  csvPrinter.print(null);
            csvPrinter.println();
            for (int i = 4; i <= orderSheet.getLastRowNum(); i++) {
                Row row = orderSheet.getRow(i);
                for (int j = 0; j <= 15; j++) {
                    if (null != row.getCell(j)) {
                        String value = row.getCell(j).toString();
            /*if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                value = value.replaceAll("\\.",",");
            }*/
                        if (j == 4) {
                            Date date = row.getCell(j).getDateCellValue();
                            DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm a");

                            value = dateFormat.format(date);
                        } else if (j == 5
                                && isEnviornmentUAT
                                && null != row.getCell(5)
                                && row.getCell(5).getCellType() != CellType.BLANK) {
                            String email = row.getCell(5).toString();
                            value = "abc".concat(email.toLowerCase(Locale.ROOT));
                        } else if (j == 14
                                || j == 15
                                || j == 1
                                && (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains("."))) {
                            value = value.split("\\.")[0];
                        }

                        csvPrinter.print(value);
                    } else csvPrinter.print(null);
                }
                // csvPrinter.print(null);
                csvPrinter.println();
            }
        }
    }

    private static void exportAddressDataToImpex(CSVPrinter csvPrinter, XSSFWorkbook addressWorkBook)
            throws IOException {

        if (addressWorkBook != null) {
            XSSFSheet addressSheet = addressWorkBook.getSheet("Address");

            Row headerRow = addressSheet.getRow(2);
            Iterator<Cell> cellIterator = headerRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (null != cell && cell.getCellType() != CellType.BLANK) {
                    csvPrinter.print(cell.toString());
                }
            }
            //  csvPrinter.print(null);
            csvPrinter.println();
            for (int i = 4; i <= addressSheet.getLastRowNum(); i++) {
                Row row = addressSheet.getRow(i);
                for (int j = 0; j <= 14; j++) {
                    if (null != row.getCell(j)) {
                        String value = row.getCell(j).toString();
            /*if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                value = value.replaceAll("\\.",",");
            }*/
                        if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                            value = value.split("\\.")[0];
                        }

                        csvPrinter.print(value);
                    } else csvPrinter.print(null);
                }
                //  csvPrinter.print(null);
                csvPrinter.println();
            }
        }
    }

    private static void createOrderEntryImpexFile(XSSFWorkbook orderEntryWorkBook) {

        CSVPrinter csvPrinter = null;
        try {

            csvPrinter =
                    new CSVPrinter(
                            new FileWriter("./Target Folder/OrderEntryImpex.impex"),
                            CSVFormat.EXCEL.withDelimiter(';').withTrim());

            if (orderEntryWorkBook != null) {
                XSSFSheet orderEntrySheet = orderEntryWorkBook.getSheet("OrderEntry");

                Cell productCatalogCell = orderEntrySheet.getRow(0).getCell(0);
                csvPrinter.print(productCatalogCell.toString());
                csvPrinter.println();

                Cell catalogVersionCell = orderEntrySheet.getRow(1).getCell(0);
                csvPrinter.print(catalogVersionCell.toString());
                csvPrinter.println();

                Row headerRow = orderEntrySheet.getRow(3);
                Iterator<Cell> cellIterator = headerRow.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (null != cell && cell.getCellType() != CellType.BLANK) {
                        csvPrinter.print(cell.toString());
                    }
                }
                //  csvPrinter.print(null);
                csvPrinter.println();
                for (int i = 5; i <= orderEntrySheet.getLastRowNum(); i++) {
                    Row row = orderEntrySheet.getRow(i);
                    for (int j = 0; j <= 7; j++) {
                        if (null != row.getCell(j)) {
                            String value = row.getCell(j).toString();
              /*if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                  value = value.replaceAll("\\.",",");
              }*/
                            if (j != 6
                                    && j != 7
                                    && row.getCell(j).getCellType() == CellType.NUMERIC
                                    && value.contains(".")) {
                                value = value.split("\\.")[0];
                            }

                            csvPrinter.print(value);
                        } else csvPrinter.print(null);
                    }
                    //  csvPrinter.print(null);
                    csvPrinter.println();
                }
            }

        } catch (Exception e) {
            System.out.println("Failed to write Order Entry Impex file to output stream : ");
            e.printStackTrace();
        } finally {
            try {
                if (csvPrinter != null) {
                    csvPrinter.flush(); // Flush and close CSVPrinter
                    csvPrinter.close();
                }
            } catch (IOException ioe) {
                System.out.println("Error when closing CSV Printer");
            }
        }
    }
}
