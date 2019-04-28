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
    public static final String EXCEL_LOCATION = "_数据来源_";
    private List<String> titles;
    private List<Map<String, String>> rows;

    public static ExcelSheetData parseFromSheet(Sheet sheet, String fileName) {
        ExcelSheetData data = new ExcelSheetData();
        data.loadDataFromSheet(sheet, fileName);
        return data;
    }

    public void loadDataFromSheet(Sheet sheet, String fileName) {
        initializeTitle(sheet.getRow(0));
        if (CollectionUtils.isEmpty(titles)) {
            return;
        }
        loadAllData(sheet, fileName);
    }

    private void loadAllData(Sheet sheet, String fileName) {
        String excelLocation = fileName + "/" + sheet.getSheetName();
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
            rowData.put(EXCEL_LOCATION, excelLocation);
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

    public void distinctByPhone() {
        Map<String, Map<String, String>> rowsMap = new HashMap<>();
        for (Map<String, String> row : rows) {
            String phone = row.get("家长电话");

            Map<String, String> existRow = rowsMap.get(phone);
            if (existRow == null) {
                rowsMap.put(phone, row);
            } else {
                mergeRow(existRow, row);

            }
        }

        rows.clear();
        rows.addAll(rowsMap.values());
    }

    private void mergeRow(Map<String, String> existRow, Map<String, String> row) {
        for (Map.Entry<String, String> entry : row.entrySet()) {
            String newValue = entry.getValue();
            String originValue = existRow.get(entry.getKey());

            if (StringUtils.isEmpty(originValue)) {
                originValue = newValue;
            } else if (!originValue.contains(newValue)) {
                originValue = originValue + ";" + newValue;
            } else {
                continue;
            }
            existRow.put(entry.getKey(), originValue);
        }
    }
}
