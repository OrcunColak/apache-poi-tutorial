package com.colak;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ReadExcelTest {

    public static void main(String[] args) {
        List<Order> orderList = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(new File("demo.xlsx"))) {
            log.info("Number of sheets: {}", workbook.getNumberOfSheets());

            workbook.forEach(sheet -> {
                log.info("Title of sheet => {}", sheet.getSheetName());

                readRow(sheet, orderList);
            });
        } catch (EncryptedDocumentException | IOException e) {
            log.error(e.getMessage(), e);
        }
    }

    private static void readRow(Sheet sheet, List<Order> orderList) {
        DataFormatter dataFormatter = new DataFormatter();
        int index = 0;
        for (Row row : sheet) {
            // Skip first row
            if (index++ == 0) {
                continue;
            }
            Order order = new Order();

            if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.NUMERIC) {
                order.setId((long) row.getCell(0).getNumericCellValue());
            }

            if (row.getCell(1) != null) {
                order.setCustomerName(dataFormatter.formatCellValue(row.getCell(1)));
            }

            Cell dateCell = row.getCell(2);
            if (DateUtil.isCellDateFormatted(dateCell)) {
                LocalDateTime localDateTime = dateCell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
                order.setOrderDate(localDateTime);
            }

            if (row.getCell(3) != null && row.getCell(3).getCellType() == CellType.NUMERIC) {
                order.setTotalAmount(row.getCell(3).getNumericCellValue());
            }
            orderList.add(order);
        }
    }
}
