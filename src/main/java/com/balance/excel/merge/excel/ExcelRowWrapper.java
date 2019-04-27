package com.balance.excel.merge.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public class ExcelRowWrapper<T> {
    private T data;
    private String sheetName;
    private int rowIndex;
}
