package com.colak;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class WriteToFileTest {

    public static void main(String[] args) throws Exception {
        // Creating a list of Order objects
        List<Order> orders = new ArrayList<>();
        orders.add(new Order(1L, "John Doe", LocalDateTime.now(), 100.0));
        orders.add(new Order(2L, "Jane Smith", LocalDateTime.now(), 150.0));
        orders.add(new Order(3L, "Alice Johnson", LocalDateTime.now(), 200.0));
        orders.add(new Order(4L, "Bob Brown", LocalDateTime.now(), 180.0));
        orders.add(new Order(5L, "Emily Davis", LocalDateTime.now(), 220.0));
        orders.add(new Order(6L, "Michael Wilson", LocalDateTime.now(), 300.0));
        generateReport(orders);
    }

    private static void generateReport(List<Order> orders) throws IllegalAccessException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Order Report");

            ExcelExporter excelExporter = new ExcelExporter(sheet,orders);
            excelExporter.export();

            // Write workbook to file
            Path path = Paths.get("order_report.xlsx");

            String filePath = path.toString();
            try (FileOutputStream fileOut = new FileOutputStream(filePath, false)) {
                workbook.write(fileOut);
                fileOut.flush();
            }
        } catch (IOException exception) {
            log.info("Exception : ", exception);
        }
    }
}
