package com.colak;


import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Order {
    @ExcelColumn(name = "ID")
    private Long id;

    @ExcelColumn(name = "CUSTOMER NAME")
    private String customerName;

    @ExcelColumn(name = "ORDER DATE")
    private LocalDateTime orderDate;

    @ExcelColumn(name = "TOTAL AMOUNT")
    private double totalAmount;
}
