package pits;

import org.apache.commons.csv.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class OrderProcessingMain {

  private static final int orderSheetColumnCount = 16;
  private static final int orderSheetIdIndex = 1;
  private static final int orderSheetEmailIndex = 5;
  private static final int orderSheetOrderDateIndex = 4;
  private static final int orderSheetDiscountIndex = 12;
  private static final int orderSheetDeliveryAddressIndex = 15;
  private static final int orderSheetPaymentAddressIndex = 16;

  private static final int addressSheetColumnCount = 15;
  private static final int addressSheetIdIndex = 1;
  private static final int addressSheetCustomerUidIndex = 3;

  private static final int orderEntrySheetColumnCount = 7;
  private static final int orderEntrySheetIdIndex = 1;
  private static final int orderEntrySheetBasePriceIndex = 6;
  private static final int orderEntrySheetTotalPriceIndex = 7;

  private static final int orderSheetStartRow = 4;
  private static final int orderEntrySheetStartRow = 5;
  private static final int addressSheetStartRow = 4;

  public static void main(String[] args) {

    try {

      Scanner scanner = new Scanner(System.in);
      boolean isEnvironmentUAT = true;

      System.out.println("Select the environment : (1/2)");
      System.out.println("1. UAT ");
      System.out.println("2. PROD");

      if (scanner.nextLine().equalsIgnoreCase("2")) {
        isEnvironmentUAT = false;
      }
      System.out.println("Importing Excel files");
      XSSFWorkbook orderWorkBook = importOrderWorkBook(scanner);

      XSSFWorkbook orderEntryWorkBook = importOrderEntryWorkBook(scanner);

      XSSFWorkbook addressWorkBook = importAddressWorkBook(scanner);

      List<CSVRecord> list = importExportCSVFile(scanner);

      long start = System.currentTimeMillis();

      XSSFWorkbook deletedEntriesWorkBook = new XSSFWorkbook();

      System.out.println("Validating Customer Email with Order Email.....");
      validateCustomerEmailInOrder(orderWorkBook, list, deletedEntriesWorkBook);

      System.out.println("Validating Order Payment and Delivery Address.....");
      validateOrderAddress(orderWorkBook, addressWorkBook, deletedEntriesWorkBook);

      System.out.println("Validating Customer Order Entry with Order Id.....");
      validateOrderEntryMapping(orderWorkBook, orderEntryWorkBook, deletedEntriesWorkBook);

      System.out.println("Started creating Order-Address impex file");
      createOrderAndAddressImpexFile(
              orderWorkBook, addressWorkBook, deletedEntriesWorkBook, isEnvironmentUAT);

      System.out.println("Started creating Order Entry impex file");
      createOrderEntryImpexFile(orderEntryWorkBook);

      System.out.println("Impex files created");

      long end = System.currentTimeMillis();

      System.out.println("Order processing took = " + (end - start) + "ms");

    } catch (Exception e) {
      e.printStackTrace();
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

  private static List<CSVRecord> importExportCSVFile(Scanner scanner) throws IOException {

    System.out.println("Enter Export Sheet Name with extension : ");
    String exportSheetName = scanner.nextLine();

    CSVParser exportCSVParser =
            new CSVParser(new FileReader("./Source Folder/" + exportSheetName), CSVFormat.DEFAULT);
    return exportCSVParser.getRecords();
  }

  private static void validateCustomerEmailInOrder(
          XSSFWorkbook orderWorkBook, List<CSVRecord> list, XSSFWorkbook deletedEntriesWorkBook) {

    XSSFSheet orderSheet = orderWorkBook.getSheet("Order");

    XSSFSheet deletedOrderEntrySheet = deletedEntriesWorkBook.createSheet("Deleted Orders");

    int deleteSheetRowNumber = 0;

    XSSFRow deletedCustomerHeaderRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber);

    deletedCustomerHeaderRow.createCell(0).setCellValue("Order Id");
    deletedCustomerHeaderRow.createCell(3).setCellValue("Reason");

    deleteSheetRowNumber++;

    HashSet<String> customerEmailSet = new HashSet<>();

    for (CSVRecord record : list) {
      customerEmailSet.add(record.get(0).substring(3));
    }

    for (int i = orderSheetStartRow; i <= orderSheet.getLastRowNum(); i++) {
      Row row = orderSheet.getRow(i);
      if (null != row.getCell(orderSheetEmailIndex)
              && !row.getCell(orderSheetEmailIndex).getStringCellValue().isEmpty()
              && row.getCell(orderSheetEmailIndex).getCellType() != CellType.BLANK) {

        if (!customerEmailSet.contains(row.getCell(orderSheetEmailIndex).getStringCellValue())) {
          System.out.println(
          "Valid Customer not found for Order(Id) = "
                  + (int) row.getCell(orderSheetIdIndex).getNumericCellValue()
                  + " for Customer Email = "
                  + row.getCell(orderSheetEmailIndex).getStringCellValue());

          XSSFRow deletedRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber++);

          deletedRow
                  .createCell(0)
                  .setCellValue(
                          (int) orderSheet.getRow(i).getCell(orderSheetIdIndex).getNumericCellValue());
          deletedRow
                  .createCell(3)
                  .setCellValue("Reason for deletion : Customer Email is not valid");

          orderSheet.shiftRows(
                  orderSheet.getRow(i).getRowNum() + 1, orderSheet.getLastRowNum() + 1, -1);
          i--;
        }
      }
    }
  }

  private static void validateOrderAddress(
          XSSFWorkbook orderWorkBook,
          XSSFWorkbook addressWorkBook,
          XSSFWorkbook deletedEntriesWorkBook) {

    XSSFSheet orderSheet = orderWorkBook.getSheet("Order");
    XSSFSheet addressSheet = addressWorkBook.getSheet("Address");

    XSSFSheet deletedOrderEntrySheet = deletedEntriesWorkBook.getSheet("Deleted Orders");

    int deleteSheetRowNumber = deletedOrderEntrySheet.getLastRowNum() + 1;

    deleteSheetRowNumber++;

    TreeMap<Integer, String> addressIdMap = new TreeMap<>();

    for (int j = addressSheetStartRow; j <= addressSheet.getLastRowNum(); j++) {
      Row row = addressSheet.getRow(j);
      if (null != row.getCell(addressSheetIdIndex)
              && row.getCell(addressSheetIdIndex).getCellType() != CellType.BLANK
              && null != row.getCell(addressSheetCustomerUidIndex)
              && row.getCell(addressSheetCustomerUidIndex).getCellType() != CellType.BLANK
              && row.getCell(addressSheetCustomerUidIndex).getCellType() != CellType.NUMERIC) {
        addressIdMap.put(
                (int) row.getCell(addressSheetIdIndex).getNumericCellValue(),
                row.getCell(addressSheetCustomerUidIndex).getStringCellValue());
      }
    }

    for (int i = orderSheetStartRow; i <= orderSheet.getLastRowNum(); i++) {
      Row row = orderSheet.getRow(i);
      if (null != row.getCell(orderSheetDeliveryAddressIndex)
              && null != row.getCell(orderSheetPaymentAddressIndex)
              && row.getCell(orderSheetDeliveryAddressIndex).getCellType() != CellType.BLANK
              && row.getCell(orderSheetPaymentAddressIndex).getCellType() != CellType.BLANK) {

        if (!addressIdMap.containsKey(
                (int) row.getCell(orderSheetDeliveryAddressIndex).getNumericCellValue())
                || !addressIdMap.containsKey(
                (int) row.getCell(orderSheetPaymentAddressIndex).getNumericCellValue())) {
          System.out.println(
                  "Payment or Delivery Address ind not available for Order(Id) = "
                          + (int) row.getCell(orderSheetIdIndex).getNumericCellValue()
                          + " for AddressId = "
                          + (int) row.getCell(orderSheetDeliveryAddressIndex).getNumericCellValue());

          XSSFRow deletedRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber++);

          deletedRow
                  .createCell(0)
                  .setCellValue(
                          (int) orderSheet.getRow(i).getCell(orderSheetIdIndex).getNumericCellValue());
          deletedRow
                  .createCell(3)
                  .setCellValue(
                          "Reason for deletion : Payment or Delivery Address not available in Address Workbook");

          // un comment below to enable Order row removing
           orderSheet.shiftRows(
                  orderSheet.getRow(i).getRowNum() + 1, orderSheet.getLastRowNum() + 1, -1);
          i--;
        }
      }
    }
  }

  private static void validateOrderEntryMapping(
          XSSFWorkbook orderWorkBook,
          XSSFWorkbook orderEntryWorkBook,
          XSSFWorkbook deletedEntriesWorkBook) {

    XSSFSheet orderSheet = orderWorkBook.getSheet("Order");
    XSSFSheet orderEntrySheet = orderEntryWorkBook.getSheet("OrderEntry");

    XSSFSheet deletedOrderEntrySheet = deletedEntriesWorkBook.createSheet("Deleted Order Entry");

    int deleteSheetRowNumber = 0;

    XSSFRow deletedCustomerHeaderRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber);

    deletedCustomerHeaderRow.createCell(0).setCellValue("Order Id");
    deletedCustomerHeaderRow.createCell(3).setCellValue("Reason");

    deleteSheetRowNumber++;

    TreeMap<Integer, String> orderMap = new TreeMap<>();

    for (int j = orderSheetStartRow; j <= orderSheet.getLastRowNum(); j++) {
      Row row = orderSheet.getRow(j);
      orderMap.put(
              (int) row.getCell(orderSheetIdIndex).getNumericCellValue(),
              row.getCell(orderSheetEmailIndex).getStringCellValue());
    }

    for (int i = orderEntrySheetStartRow; i <= orderEntrySheet.getLastRowNum(); i++) {
      Row row = orderEntrySheet.getRow(i);
      if (!orderMap.containsKey((int) row.getCell(orderSheetIdIndex).getNumericCellValue())) {
         System.out.println(
        "Order not available for Order Entry with Id : "
                + (int) row.getCell(orderSheetIdIndex).getNumericCellValue());

        XSSFRow deletedRow = deletedOrderEntrySheet.createRow(deleteSheetRowNumber++);

        deletedRow
                .createCell(0)
                .setCellValue(
                        (int) orderEntrySheet.getRow(i).getCell(orderSheetIdIndex).getNumericCellValue());
        deletedRow.createCell(3).setCellValue("Reason for deletion : No Mapping in Order Sheet");

        orderEntrySheet.shiftRows(
                orderEntrySheet.getRow(i).getRowNum() + 1, orderEntrySheet.getLastRowNum() + 1, -1);
        i--;
      }
    }
  }

  private static void createOrderAndAddressImpexFile(
          XSSFWorkbook orderWorkbook,
          XSSFWorkbook addressWorkBook,
          XSSFWorkbook deletedEntriesWorkBook,
          boolean isEnvironmentUAT) {

    CSVPrinter csvPrinter = null;
    try {

      csvPrinter =
              new CSVPrinter(
                      new FileWriter("./Target Folder/OrderImpex.impex"),
                      CSVFormat.EXCEL.withDelimiter(';').withTrim().withQuoteMode(QuoteMode.MINIMAL));

      exportOrderDataToImpex(csvPrinter, orderWorkbook, isEnvironmentUAT);

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
          CSVPrinter csvPrinter, XSSFWorkbook orderWorkbook, boolean isEnvironmentUAT)
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
      for (int i = orderSheetStartRow; i <= orderSheet.getLastRowNum(); i++) {
        Row row = orderSheet.getRow(i);
        if (null != row.getCell(orderSheetIdIndex)
                && row.getCell(orderSheetIdIndex).getCellType() != CellType.BLANK) {
          for (int j = 0; j <= orderSheetColumnCount; j++) {
            if (j != orderSheetDiscountIndex
                    && null != row.getCell(j)
                    && row.getCell(j).getCellType() != CellType.BLANK) {
              String value = row.getCell(j).toString();

              if (j == orderSheetOrderDateIndex
                      && null != row.getCell(orderSheetOrderDateIndex)
                      && row.getCell(orderSheetOrderDateIndex).getCellType() != CellType.BLANK) {
                Date date = row.getCell(j).getDateCellValue();
                DateFormat dateFormat = new SimpleDateFormat("dd/MM/yy HH:mm");

                value = dateFormat.format(date);
              } else if (j == orderSheetEmailIndex
                      && null != row.getCell(orderSheetEmailIndex)
                      && row.getCell(orderSheetEmailIndex).getCellType() != CellType.BLANK) {
                String email = row.getCell(orderSheetEmailIndex).toString();
                if (isEnvironmentUAT) value = "abc".concat(email.toLowerCase(Locale.ROOT));
                else value = email.toLowerCase(Locale.ROOT);
              } else if (j == orderSheetDeliveryAddressIndex
                      || j == orderSheetPaymentAddressIndex
                      || j == orderSheetIdIndex
                      && (row.getCell(j).getCellType() == CellType.NUMERIC
                      && value.contains("."))) {
                value = value.split("\\.")[0];
              }

              csvPrinter.print(value);
            } else if (j != orderSheetDiscountIndex) csvPrinter.print(null);
          }
          csvPrinter.println();
        }
      }
    }
  }

  private static void exportAddressDataToImpex(
          CSVPrinter csvPrinter, XSSFWorkbook addressWorkBook, XSSFWorkbook deletedEntriesWorkBook)
          throws IOException {

    if (addressWorkBook != null) {
      Map<String, String> orderAddressMap = new HashMap<>();
      XSSFSheet addressSheet = addressWorkBook.getSheet("Address");
      XSSFSheet deletedAddressSheet = deletedEntriesWorkBook.createSheet("Deleted Addresses");

      int deleteSheetRowNumber = 0;

      XSSFRow deletedCustomerHeaderRow = deletedAddressSheet.createRow(deleteSheetRowNumber);

      deletedCustomerHeaderRow.createCell(0).setCellValue("Address Id");
      deletedCustomerHeaderRow.createCell(1).setCellValue("Customer Uid");
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
      for (int i = addressSheetStartRow; i <= addressSheet.getLastRowNum(); i++) {
        Row row = addressSheet.getRow(i);
        if (null != row.getCell(addressSheetIdIndex)
                && row.getCell(addressSheetIdIndex).getCellType() != CellType.BLANK
                && row.getCell(addressSheetCustomerUidIndex).getCellType() != CellType.NUMERIC) {
          if (!orderAddressMap.containsKey(row.getCell(addressSheetIdIndex).toString())) {
            orderAddressMap.put(
                    row.getCell(addressSheetIdIndex).toString(),
                    row.getCell(addressSheetCustomerUidIndex).toString());
            for (int j = 0; j <= addressSheetColumnCount; j++) {
              if (null != row.getCell(j) && row.getCell(j).getCellType() != CellType.BLANK) {
                String value = row.getCell(j).toString();

                if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                  value = value.split("\\.")[0];
                }
                csvPrinter.print(value);
              } else csvPrinter.print(null);
            }
            csvPrinter.println();
          } else {
            XSSFRow deletedRow = deletedAddressSheet.createRow(deleteSheetRowNumber++);

            deletedRow
                    .createCell(0)
                    .setCellValue(addressSheet.getRow(i).getCell(addressSheetIdIndex).toString());
            deletedRow
                    .createCell(1)
                    .setCellValue(
                            addressSheet.getRow(i).getCell(addressSheetCustomerUidIndex).toString());
            deletedRow.createCell(3).setCellValue("Reason for deletion : Duplicate AddressId");
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
        for (int i = orderEntrySheetStartRow; i <= orderEntrySheet.getLastRowNum(); i++) {
          Row row = orderEntrySheet.getRow(i);
          if (null != row.getCell(orderEntrySheetIdIndex)
                  && row.getCell(orderEntrySheetIdIndex).getCellType() != CellType.BLANK) {
            for (int j = 0; j <= orderEntrySheetColumnCount; j++) {
              if (null != row.getCell(j) && row.getCell(j).getCellType() != CellType.BLANK) {
                String value = row.getCell(j).toString();

                if (j != orderEntrySheetBasePriceIndex
                        && j != orderEntrySheetTotalPriceIndex
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
