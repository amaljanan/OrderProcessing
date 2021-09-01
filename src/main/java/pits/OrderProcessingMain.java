package pits;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

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
      System.out.println("Importing Excel files");
      XSSFWorkbook orderWorkBook = importOrderWorkBook(scanner);

      XSSFWorkbook orderEntryWorkBook = importOrderEntryWorkBook(scanner);

      XSSFWorkbook addressWorkBook = importAddressWorkBook(scanner);

      XSSFWorkbook deletedEntriesWorkBook = new XSSFWorkbook();
      
      validateOrderEntryMapping(orderWorkBook,orderEntryWorkBook,deletedEntriesWorkBook);

      System.out.println("Started creating Order-Address impex file");
      createOrderAndAddressImpexFile(orderWorkBook, addressWorkBook, deletedEntriesWorkBook, isEnviornmentUAT);

      System.out.println("Started creating Order Entry impex file");
      createOrderEntryImpexFile(orderEntryWorkBook);

      System.out.println("Impex files created");
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void validateOrderEntryMapping(XSSFWorkbook orderWorkBook, XSSFWorkbook orderEntryWorkBook, XSSFWorkbook deletedEntriesWorkBook) {

    XSSFSheet orderSheet = orderWorkBook.getSheet("Order");
    XSSFSheet orderEntrySheet = orderEntryWorkBook.getSheet("OrderEntry");

    XSSFSheet deletedOrderEntrySheet = deletedEntriesWorkBook.createSheet("Deleted Order Entry");

    int deleteSheetRowNumber = 0;

    XSSFRow deletedCustomerHeaderRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber);

    deletedCustomerHeaderRow.createCell(0).setCellValue("Order Id");
    deletedCustomerHeaderRow.createCell(3).setCellValue("Reason");

    deleteSheetRowNumber++;

    TreeMap<Integer,String> orderMap = new TreeMap<>();

    for (int j = 4; j <= orderSheet.getLastRowNum(); j++) {
      Row row = orderSheet.getRow(j);
        orderMap.put((int) row.getCell(1).getNumericCellValue(),row.getCell(5).getStringCellValue());
    }

    for (int i = 5; i <= orderEntrySheet.getLastRowNum(); i++) {
      Row row = orderEntrySheet.getRow(i);
        if(!orderMap.containsKey((int) row.getCell(1).getNumericCellValue())) {
          System.out.println("Order not available for Order Entry with Id : "+(int) row.getCell(1).getNumericCellValue());

          XSSFRow deletedRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber++);

          deletedRow
                  .createCell(0)
                  .setCellValue((int) orderEntrySheet.getRow(i).getCell(1).getNumericCellValue());
          deletedRow
                  .createCell(3)
                  .setCellValue("Reason for deletion : No Mapping in Order Sheet");

          orderEntrySheet.shiftRows(
                  orderEntrySheet.getRow(i).getRowNum() + 1, orderEntrySheet.getLastRowNum() + 1, -1);
          i--;
        }
    }
  }

  private static XSSFWorkbook importOrderWorkBook(Scanner scanner) throws IOException {

    System.out.println("Enter Order Workbook name with extension : ");
    String fileName = scanner.nextLine();

    FileInputStream fileInputStream = new FileInputStream("./Source Folder/" + fileName);

    return new XSSFWorkbook(fileInputStream);
  }

  private static XSSFWorkbook importOrderEntryWorkBook(Scanner scanner) throws IOException {

    System.out.println("Enter Order Entry Workbook name with extension : ");
    String fileName = scanner.nextLine();

    FileInputStream fileInputStream = new FileInputStream("./Source Folder/" + fileName);

    return new XSSFWorkbook(fileInputStream);
  }

  private static XSSFWorkbook importAddressWorkBook(Scanner scanner) throws IOException {

    System.out.println("Enter Address Workbook name with extension : ");
    String fileName = scanner.nextLine();

    FileInputStream fileInputStream = new FileInputStream("./Source Folder/" + fileName);

    return new XSSFWorkbook(fileInputStream);
  }

  private static void createOrderAndAddressImpexFile(
          XSSFWorkbook orderWorkbook, XSSFWorkbook addressWorkBook, XSSFWorkbook deletedEntriesWorkBook, boolean isEnviornmentUAT) {

    CSVPrinter csvPrinter = null;
    try {

      csvPrinter =
              new CSVPrinter(
                      new FileWriter("./Target Folder/OrderImpex.impex"),
                      CSVFormat.EXCEL.withDelimiter(';').withTrim().withQuoteMode(QuoteMode.MINIMAL));

      exportOrderDataToImpex(csvPrinter, orderWorkbook, isEnviornmentUAT);

      csvPrinter.println();

      exportAddressDataToImpex(csvPrinter, addressWorkBook, deletedEntriesWorkBook);

    } catch (Exception e) {
      System.out.println("Failed to write Order and Address Impex file to output stream : ");
      e.printStackTrace();
    } finally {
      try {
        if (csvPrinter != null) {
          csvPrinter.flush();
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
      csvPrinter.println();
      for (int i = 4; i <= orderSheet.getLastRowNum(); i++) {
        Row row = orderSheet.getRow(i);
        if (null != row.getCell(1) && row.getCell(1).getCellType() != CellType.BLANK) {
          for (int j = 0; j <= 15; j++) {
            if (null != row.getCell(j) && row.getCell(j).getCellType() != CellType.BLANK) {
              String value = row.getCell(j).toString();

              if (j == 4
                      && null != row.getCell(4)
                      && row.getCell(4).getCellType() != CellType.BLANK) {
                Date date = row.getCell(j).getDateCellValue();
                DateFormat dateFormat = new SimpleDateFormat("dd/MM/yy HH:mm");

                value = dateFormat.format(date);
              } else if (j == 5
                      && null != row.getCell(5)
                      && row.getCell(5).getCellType() != CellType.BLANK) {
                String email = row.getCell(5).toString();
                if (isEnviornmentUAT) value = "abc".concat(email.toLowerCase(Locale.ROOT));
                else value = email.toLowerCase(Locale.ROOT);
              } else if (j == 14
                      || j == 15
                      || j == 1
                      && (row.getCell(j).getCellType() == CellType.NUMERIC
                      && value.contains("."))) {
                value = value.split("\\.")[0];
              }

              csvPrinter.print(value);
            } else csvPrinter.print(null);
          }
          csvPrinter.println();
        }
      }
    }
  }

  private static void exportAddressDataToImpex(CSVPrinter csvPrinter, XSSFWorkbook addressWorkBook, XSSFWorkbook deletedEntriesWorkBook)
          throws IOException {

    if (addressWorkBook != null) {
      Map<String, String> orderAddressMap = new HashMap<>();
      XSSFSheet addressSheet = addressWorkBook.getSheet("Address");
      XSSFSheet deletedAddressSheet = deletedEntriesWorkBook.createSheet("Deleted Addresses");


      int deleteSheetRowNumber = 0;

      XSSFRow deletedCustomerHeaderRow = deletedAddressSheet.createRow(deleteSheetRowNumber);

      deletedCustomerHeaderRow.createCell(0).setCellValue("Address Id");
      deletedCustomerHeaderRow.createCell(1).setCellValue("Order Id");
      deletedCustomerHeaderRow.createCell(3).setCellValue("Reason");

      deleteSheetRowNumber++;


      Row headerRow = addressSheet.getRow(2);
      Iterator<Cell> cellIterator = headerRow.cellIterator();
      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        if (null != cell && cell.getCellType() != CellType.BLANK) {
          csvPrinter.print(cell.toString());
        }
      }
      csvPrinter.println();
      for (int i = 4; i <= addressSheet.getLastRowNum(); i++) {
        Row row = addressSheet.getRow(i);
        if (null != row.getCell(1) && row.getCell(1).getCellType() != CellType.BLANK) {
          if(!orderAddressMap.containsKey(row.getCell(1).toString())){
            orderAddressMap.put(row.getCell(1).toString(),row.getCell(2).toString());
            for (int j = 0; j <= 14; j++) {
              if (null != row.getCell(j) && row.getCell(j).getCellType() != CellType.BLANK) {
                String value = row.getCell(j).toString();

                if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                  value = value.split("\\.")[0];
                }
                csvPrinter.print(value);
              } else csvPrinter.print(null);
            }
            csvPrinter.println();
          }else{
            XSSFRow deletedRow = deletedAddressSheet.createRow(deleteSheetRowNumber++);

            deletedRow
                    .createCell(0)
                    .setCellValue(addressSheet.getRow(i).getCell(1).toString());
            deletedRow
                    .createCell(1)
                    .setCellValue(addressSheet.getRow(i).getCell(2).toString());
            deletedRow
                    .createCell(3)
                    .setCellValue("Reason for deletion : Duplicate AddressId");
          }
        }
      }

      FileOutputStream deletedRecordsFileOutputStream =
              new FileOutputStream("./Target Folder/DeletedRecords.xlsx");
      deletedEntriesWorkBook.write(deletedRecordsFileOutputStream);
      deletedRecordsFileOutputStream.close();
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
        csvPrinter.println();
        for (int i = 5; i <= orderEntrySheet.getLastRowNum(); i++) {
          Row row = orderEntrySheet.getRow(i);
          if (null != row.getCell(1) && row.getCell(1).getCellType() != CellType.BLANK) {
            for (int j = 0; j <= 7; j++) {
              if (null != row.getCell(j) && row.getCell(j).getCellType() != CellType.BLANK) {
                String value = row.getCell(j).toString();

                if (j != 6
                        && j != 7
                        && row.getCell(j).getCellType() == CellType.NUMERIC
                        && value.contains(".")) {
                  value = value.split("\\.")[0];
                }

                csvPrinter.print(value);
              } else csvPrinter.print(null);
            }
            csvPrinter.println();
          }
        }
      }

    } catch (Exception e) {
      System.out.println("Failed to write Order Entry Impex file to output stream : ");
      e.printStackTrace();
    } finally {
      try {
        if (csvPrinter != null) {
          csvPrinter.flush();
          csvPrinter.close();
        }
      } catch (IOException ioe) {
        System.out.println("Error when closing CSV Printer");
      }
    }
  }
}
