package pits;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Scanner;

public class OrderProcessingMain {

    public static void main(String[] args) {

        try {

            Scanner scanner = new Scanner(System.in);
            boolean isEnviornmentUAT = true;

            XSSFWorkbook customerWorkbook = importCustomerWorkBook(scanner);


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static XSSFWorkbook importCustomerWorkBook (Scanner scanner) throws IOException {

        System.out.println("Enter Order Workbook name with extension : ");
        String customerSheetName = scanner.nextLine();

        FileInputStream customerFileInputStream = new FileInputStream("./Source/" + customerSheetName);
        // customerFileInputStream.close();

        return new XSSFWorkbook(customerFileInputStream);
    }
}
