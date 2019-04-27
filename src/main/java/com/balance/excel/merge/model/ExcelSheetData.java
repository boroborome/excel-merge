package com.balance.excel.merge.model;

import lombok.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelSheetData {
    private List<String> titles;
    private List<Map<String, String>> rows;

    public static ExcelSheetData parseFromSheet(Sheet sheet) {
        ExcelSheetData data = new ExcelSheetData();
        data.loadDataFromSheet(sheet);
        return data;
    }

    public void loadDataFromSheet(Sheet sheet) {
        initializeTitle(sheet.getRow(0));
        if (CollectionUtils.isEmpty(titles)) {
            return;
        }
        loadAllData(sheet);
    }

    private void loadAllData(Sheet sheet) {
        rows = new ArrayList<>();
        for (int rowIndex = 1; rowIndex < sheet.getLastRowNum() ; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                break;
            }
            Map<String, String> rowData = new HashMap<>();

            for (int column = 0; column < titles.size(); column++) {
                Cell c = row.getCell(column);
                if (c == null) {
                    continue;
                }
                c.setCellType(CellType.STRING);
                String value = c.getStringCellValue().trim();
                rowData.put(titles.get(column), value);
            }

            if (CollectionUtils.isEmpty(rowData)) {
                break;
            }
            rows.add(rowData);
        }
    }


    private void initializeTitle(Row row) {
        if (row == null) {
            return;
        }

        titles = new ArrayList<>();
        for (int i = 0; ; i++) {
            Cell c = row.getCell(i);
            if (c == null) {
                break;
            }
            c.setCellType(CellType.STRING);
            String value = c.getStringCellValue().trim();
            if (StringUtils.isEmpty(value)) {
                break;
            }

            titles.add(value);
        }
    }

    public ExcelSheetData merge(ExcelSheetData otherData) {
        if (this.titles == null) {
            titles = new ArrayList<>();
        }
        if (rows == null) {
            rows = new ArrayList<>();
        }

        for (String title : otherData.getTitles()) {
            if (titles.contains(title)) {
                continue;
            }
            titles.add(title);
        }

        rows.addAll(otherData.getRows());
        return this;
    }
}
