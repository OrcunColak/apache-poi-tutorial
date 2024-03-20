package com.colak;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.List;

public class ExcelExporter {
    private final Sheet sheet;
    private final List<?> objects;
    private Field[] fields;

    public ExcelExporter(Sheet sheet, List<?> objects) {
        this.sheet = sheet;
        this.objects = objects;
    }

    public void export() throws IllegalAccessException {
        createHeaderRow();
        populateDataRows();
    }

    private void createHeaderRow() {
        Object object = objects.getFirst();
        // Create header row
        Row headerRow = sheet.createRow(0);
        fields = object.getClass().getDeclaredFields();
        // Create header
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(column.name());
            }
        }
    }

    private void populateDataRows() throws IllegalAccessException {
        int rowNum = 1;
        for (Object object : objects) {
            Row row = sheet.createRow(rowNum++);

            for (int fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) {
                Field field = fields[fieldIndex];
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);
                    Object value = field.get(object);
                    if (value != null) {
                        Cell cell = row.createCell(fieldIndex);
                        cell.setCellValue(value.toString());
                    }
                    field.setAccessible(false);
                }
            }
        }
    }
}
