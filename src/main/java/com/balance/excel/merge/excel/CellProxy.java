package com.balance.excel.merge.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

@Slf4j
public class CellProxy {

    private final Cell cell;

    private CellProxy(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        this.cell = cell;
    }

    public CellProxy value(String value) {
        cell.setCellValue(value);
        return this;
    }

    public CellProxy cellStyle(CellStyle style) {
        cell.setCellStyle(style);
        return this;
    }

    public CellProxy merge(int rowSize, int columnSize) {
        if (columnSize > 1) {
            cell.getSheet().addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex() + rowSize - 1,
                    cell.getColumnIndex(), cell.getColumnIndex() + columnSize - 1));
        }
        return this;
    }

    public static CellProxy of(Sheet sheet, int row, int column) {
        return new CellProxy(sheet, row, column);
    }

    public static void write(Sheet sheet, int rowIndex, int columnIndex, CellStyle style, String... texts) {
        int column = columnIndex;
        for (String text : texts) {
            CellProxy cellProxy = CellProxy.of(sheet, rowIndex, column)
                    .value(text);
            if (style != null) {
                cellProxy.cellStyle(style);
            }
            ++column;
        }
    }

    private static Cell getCell(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    public static void write(Row row, int columnIndex, String value) {
        Cell cell = getCell(row, columnIndex);
        cell.setCellValue(value);
    }

    public static void write(Row row, int columnIndex, int value) {
        Cell cell = getCell(row, columnIndex);
        cell.setCellValue(value);
    }

    public void value(Date date, SimpleDateFormat dateFormat, CellStyle dateCellStyle) {
        Calendar calendar = Calendar.getInstance(dateFormat.getTimeZone());
        try {
            Date roundDate = dateFormat.parse(dateFormat.format(date));
            calendar.setTime(roundDate);
            cell.setCellValue(calendar);
            cell.setCellStyle(dateCellStyle);
        } catch (ParseException e) {
            log.error("Failed to parse date", e);
        }
    }

    public void value(double value) {
        cell.setCellValue(value);
    }

    public int getColumnSize() {
        for (CellRangeAddress range : cell.getSheet().getMergedRegions()) {
            if (range.containsColumn(cell.getColumnIndex()) && range.containsRow(cell.getRowIndex())) {
                return range.getLastColumn() - range.getFirstColumn() + 1;
            }
        }
        return 1;
    }
}
