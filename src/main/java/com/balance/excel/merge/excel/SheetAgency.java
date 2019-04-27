package com.balance.excel.merge.excel;

import com.balance.excel.merge.util.MapBuilder;
import com.balance.excel.merge.util.convert.primitive.PrimitiveConverter;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.StringUtils;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Slf4j
public class SheetAgency {
    private static final Function<Cell, Object> NULLABLE_NUMBER_READER_FUNCTION = c -> {
        if (CellType.BLANK.equals(c.getCellTypeEnum())) {
            return null;
        }
        return c.getNumericCellValue();
    };

    private static final Map<Class, Function<Cell, Object>> CellReaderMap = MapBuilder
            .<Class, Function<Cell, Object>>of(String.class, (c) -> {
                c.setCellType(CellType.STRING);
                return c.getStringCellValue().trim();
            })
            .and(double.class, c -> c.getNumericCellValue())
            .and(Double.class, NULLABLE_NUMBER_READER_FUNCTION)
            .and(Integer.class, c -> (int) c.getNumericCellValue())
            .and(Long.class, c -> (long) c.getNumericCellValue())
            .and(Date.class, c -> c.getDateCellValue())
            .build();

    private final Sheet sheet;
    private SimpleDateFormat dateFormat;
    private CellStyle defaultDateCellStyle;

    private Map<String, CellStyle> cellStyleMap = new HashMap<>();
    private Map<String, Short> dataFormatMap = new HashMap<>();
    private CellStyle defaultNumberCellStyle;

    private int row;
    private int column;

    public SheetAgency(Sheet sheet) {
        this.sheet = sheet;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public SheetAgency datePattern(String datePattern, TimeZone zone) {
        dateFormat = new SimpleDateFormat(datePattern);
        if (zone != null) {
            dateFormat.setTimeZone(zone);
        }

        cellStyle(datePattern, ExcelCellStyle.builder()
                .dataFormat(datePattern)
                .build());
        defaultDateCellStyle = cellStyleMap.get(datePattern);

        return this;
    }

    public SheetAgency defaultNumberCellStyle(String styleName) {
        defaultNumberCellStyle = cellStyleMap.get(styleName);
        return this;
    }

    public SheetAgency defaultNumberCellStyleWithDataFormat(String numFormat) {
        cellStyle(numFormat, ExcelCellStyle.builder()
                .dataFormat(numFormat)
                .build());

        return defaultNumberCellStyle(numFormat);
    }

    public SheetAgency columnStyle(String styleName, int... columnIndexes) {
        CellStyle cellStyle = cellStyleMap.get(styleName);
        for (int index : columnIndexes) {
            sheet.setDefaultColumnStyle(index, cellStyle);
        }
        return this;
    }

    public SheetAgency write(Object... values) {
        return writeWithStyle((String) null, values);
    }

    public CellStyle getColumnStyle(int column) {
        CellStyle columnStyle = sheet.getColumnStyle(column);
        return columnStyle.getIndex() == 0 ? null : columnStyle;
    }

    public SheetAgency writeWithStyle(String styleName, Object... values) {
        CellStyle assignedStyle = StringUtils.isEmpty(styleName)
            ? null
            : cellStyleMap.get(styleName);
        for (Object v : values) {
            CellProxy cellProxy = CellProxy.of(sheet, row, column);
            CellStyle columnStyle = getColumnStyle(column);
            if (v instanceof Date) {
                Date date = (Date) v;
                CellStyle cellStyle = assignedStyle != null ? assignedStyle
                        : (columnStyle != null ? columnStyle : defaultDateCellStyle);
                cellProxy.value(date, dateFormat, cellStyle);
            } else if (v instanceof Number) {
                CellStyle cellStyle = assignedStyle != null ? assignedStyle
                        : (columnStyle != null ? columnStyle : defaultNumberCellStyle);
                if (cellStyle != null) {
                    cellProxy.cellStyle(cellStyle);
                }
                cellProxy.value(((Number) v).doubleValue());
            } else {
                if (assignedStyle != null) {
                    cellProxy.cellStyle(assignedStyle);
                }
                if (v instanceof List) {
                    List<String> strList = (List<String>) ((List) v).stream()
                            .map(item -> String.valueOf(item))
                            .collect(Collectors.toList());
                    cellProxy.value(String.join(",", strList));
                } else if (v == null) {
                    // not need to write
//                cellProxy.value("");
                } else {
                    cellProxy.value(String.valueOf(v));
                }
            }

            column += cellProxy.getColumnSize();
        }
        return this;
    }

    private CellStyle translateExcelCell(ExcelCellStyle excelCellStyle) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        applyCellStyle(cellStyle, excelCellStyle);
        return cellStyle;
    }

    public <T> SheetAgency writeWithStyle(ExcelObjectDefinition<T> objectDefinition, Object... values) {
        List<ExcelColumnDefinition> fields = objectDefinition.getFields();
        for (int i = 0; i < values.length; ++i) {
            Object v = values[i];
            ExcelCellStyle assignedStyle = fields.get(i).getCellStyle();
            CellProxy cellProxy = CellProxy.of(sheet, row, column);
            CellStyle columnStyle = getColumnStyle(column);
            if (v instanceof Date) {
                Date date = (Date) v;
                CellStyle cellStyle = assignedStyle != null ? translateExcelCell(assignedStyle)
                        : (columnStyle != null ? columnStyle : defaultDateCellStyle);
                cellProxy.value(date, dateFormat, cellStyle);
            } else if (v instanceof Number) {
                CellStyle cellStyle = assignedStyle != null ? translateExcelCell(assignedStyle)
                        : (columnStyle != null ? columnStyle : defaultNumberCellStyle);
                if (cellStyle != null) {
                    cellProxy.cellStyle(cellStyle);
                }
                cellProxy.value(((Number) v).doubleValue());
            } else {
                if (assignedStyle != null) {
                    cellProxy.cellStyle(translateExcelCell(assignedStyle));
                }
                if (v instanceof List) {
                    List<String> strList = (List<String>) ((List) v).stream()
                            .map(item -> String.valueOf(item))
                            .collect(Collectors.toList());
                    cellProxy.value(String.join(",", strList));
                } else if (v == null) {
                    // not need to write
//                cellProxy.value("");
                } else {
                    cellProxy.value(String.valueOf(v));
                }
            }

            column += cellProxy.getColumnSize();
        }
        return this;
    }

    public <T> T readCellValue(int rowIndex, int columnIndex, Class<T> dataType) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }

        Function<Cell, Object> reader = CellReaderMap.get(dataType);
        if (reader != null) {
            try {
                return (T) reader.apply(cell);
            } catch (Exception e) {
                log.error("Failed to read cell value at row:{}, column:{}, dataType:{}.",
                        rowIndex, columnIndex, dataType);
                throw e;
            }
        }
        cell.setCellType(CellType.STRING);
        String strValue = cell.getStringCellValue().trim();
        Object newValue = PrimitiveConverter.getInstance().convert(strValue, dataType);
        return (T) newValue;
    }
    public String readCellValue(int rowIndex, int columnIndex) {
        return readCellValue(rowIndex, columnIndex, String.class);
    }

    public SheetAgency locate(int row, int column) {
        if (row < 0 || column < 0) {
            throw new IllegalArgumentException("Row and Column in excel can not small then 0.");
        }
        this.row = row;
        this.column = column;
        return this;
    }

    public SheetAgency newLine() {
        row++;
        column = 0;
        return this;
    }

    public static SheetAgency of(Sheet sheet) {
        return new SheetAgency(sheet);
    }
    public static SheetAgency of(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        return new SheetAgency(sheet);
    }

    public SheetAgency mergeCell(int rowSize, int columnSize) {
        if (columnSize > 1 || rowSize > 1) {
            sheet.addMergedRegion(new CellRangeAddress(row, row + rowSize - 1,
                    column, column + columnSize - 1));
        }

        return this;
    }

    public SheetAgency nextColumn() {
        CellProxy cellProxy = CellProxy.of(sheet, row, column);
        column += cellProxy.getColumnSize();
        return this;
    }

    public short getDataFormat(String dataFormat) {
        Short formatId = dataFormatMap.get(dataFormat);
        if (formatId == null) {
            short df = sheet.getWorkbook().createDataFormat().getFormat(dataFormat);
            formatId = df;
            dataFormatMap.put(dataFormat, formatId);
        }
        return formatId;
    }

    public SheetAgency columnIndex(int columnIndex) {
        return locate(row, 1);
    }

    public SheetAgency cellStyle(String styleName, ExcelCellStyle excelCellStyle) {
        CellStyle cellStyle = cellStyleMap.get(styleName);
        if (cellStyle == null) {
            cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyleMap.put(styleName, cellStyle);
        }

        applyCellStyle(cellStyle, excelCellStyle);

        return this;
    }

    private void applyCellStyle(CellStyle cellStyle, ExcelCellStyle excelCellStyle) {
        if (StringUtils.hasText(excelCellStyle.getDataFormat())) {
            cellStyle.setDataFormat(getDataFormat(excelCellStyle.getDataFormat()));
        }

        if (excelCellStyle.getFillForegroundColor() > 0) {
            cellStyle.setFillForegroundColor(excelCellStyle.getFillForegroundColor());
        }

        if (excelCellStyle.getFillPattern() != null) {
            cellStyle.setFillPattern(excelCellStyle.getFillPattern());
        }

        BorderStyle borderStyle = excelCellStyle.getBorderStyle();
        if (borderStyle != null) {
            cellStyle.setBorderBottom(borderStyle);
            cellStyle.setBorderTop(borderStyle);
            cellStyle.setBorderLeft(borderStyle);
            cellStyle.setBorderRight(borderStyle);
        }

        if (excelCellStyle.getAlignment() != null) {
            cellStyle.setAlignment(excelCellStyle.getAlignment());
        }
    }

    @AllArgsConstructor
    public class CellStyleBuilder {
        private CellStyle cellStyle;

        public CellStyleBuilder dataFormat(String format) {
            cellStyle.setDataFormat(getDataFormat(format));
            return this;
        }

        public CellStyleBuilder fillForegroundColor(short color) {
            cellStyle.setFillForegroundColor(color);
            return this;
        }

        public CellStyleBuilder fillPattern(FillPatternType fp) {
            cellStyle.setFillPattern(fp);
            return this;
        }

        public CellStyleBuilder alignment(HorizontalAlignment align) {
            cellStyle.setAlignment(align);
            return this;
        }

        public CellStyleBuilder border(BorderStyle border) {
            cellStyle.setBorderBottom(border);
            cellStyle.setBorderTop(border);
            cellStyle.setBorderLeft(border);
            cellStyle.setBorderRight(border);
            return this;
        }

        public CellStyle cellStyle() {
            return cellStyle;
        }

        public SheetAgency build() {
            return SheetAgency.this;
        }
    }
}
