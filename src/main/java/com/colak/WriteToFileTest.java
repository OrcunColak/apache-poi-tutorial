package com.colak;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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

    public static void main(String[] args) {
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

    private static void generateReport(List<Order> orders) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Order Report");

            // Create header row
            Row headerRow = sheet.createRow(0);
            String[] columns = {"Order ID", "Customer Name", "Order Date", "Total Amount"};
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }

            // Populate data rows
            int rowNum = 1;
            for (Order order : orders) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(order.getId());
                row.createCell(1).setCellValue(order.getCustomerName());
                row.createCell(2).setCellValue(order.getOrderDate().toString());
                row.createCell(3).setCellValue(order.getTotalAmount());
            }

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
