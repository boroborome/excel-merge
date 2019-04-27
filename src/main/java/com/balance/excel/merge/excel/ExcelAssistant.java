package com.balance.excel.merge.excel;

import com.balance.excel.merge.util.MessageRecorder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.TimeZone;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Slf4j
public class ExcelAssistant {
    private Workbook workbook;
    private String datePattern;
    private TimeZone zone;
    private String numFormat;

    public ExcelAssistant(Workbook workbook) {
        this.workbook = workbook;
    }

    public <T> List<T> loadAllSheet(Class<T> dataType) {
        return loadAllSheet(dataType, w -> w.getData());
    }

    public <T> List<T> loadAllSheet(Class<T> dataType, Function<ExcelRowWrapper<T>, T> checker) {
        return loadAllSheet(dataType, checker, new MessageRecorder());
    }

    public <T> List<T> loadAllSheet(Class<T> dataType, Function<ExcelRowWrapper<T>, T> checker, MessageRecorder messageRecorder) {
        ExcelObjectDefinition<T> objDefinition = ExcelObjectDefinition.from(dataType);

        List<T> allData = new ArrayList<>();
        for (int index = 0, numberOfSheets = workbook.getNumberOfSheets(); index < numberOfSheets; index++) {
            Sheet sheet = workbook.getSheetAt(index);
            ExcelRowIterator<T> dataIt = ExcelRowIterator.from(sheet, objDefinition, messageRecorder);
            dataIt.forEachRemaining(w ->
                    {
                        T data = checker.apply(w);
                        if (data != null) {
                            allData.add(w.getData());
                        }
                    });
        }
        return allData;
    }

    public static ExcelAssistant from(Workbook workbook) {
        return new ExcelAssistant(workbook);
    }

    public static ExcelAssistant from(InputStream excelStream) {
        XSSFWorkbook xssfWorkbook;
        try {
            xssfWorkbook = new XSSFWorkbook(excelStream);
        } catch (IOException e) {
            String message = "Failed in reading excel file.";
            log.error(message, e);
            return null;
        }
        return new ExcelAssistant(xssfWorkbook);
    }

    public ExcelAssistant datePattern(String datePattern, TimeZone zone) {
        this.datePattern = datePattern;
        this.zone = zone;
        return this;
    }

    public ExcelAssistant numberFormat(String numFormat) {
        this.numFormat = numFormat;
        return this;
    }

    public <T> ExcelAssistant saveToSheet(List<T> datas, String sheetName, Class<T> dataType) {
        return saveToSheet(datas.stream(), sheetName, dataType);
    }

    public <T> ExcelAssistant saveToSheet(Stream<T> datas, String sheetName, Class<T> dataType) {
        ExcelObjectDefinition<T> objectDefinition = ExcelObjectDefinition.from(dataType);
        List<String> titles = objectDefinition.getFields().stream()
                .map(c -> c.getExcelTitle())
                .collect(Collectors.toList());
        SheetAgency sheetAgency = SheetAgency.of(workbook, sheetName);
        if (StringUtils.hasText(datePattern)) {
            sheetAgency.datePattern(datePattern, zone);
        }
        if (StringUtils.hasText(numFormat)) {
            sheetAgency.defaultNumberCellStyleWithDataFormat(numFormat);
        }
        sheetAgency.write(titles.toArray());
        datas.forEach(data ->{
            List<Object> values = objectDefinition.getFields().stream()
                    .map(c -> c.getAccessor().getValue(data))
                    .collect(Collectors.toList());
            sheetAgency.newLine().writeWithStyle(objectDefinition, values.toArray());
        });
        return this;
    }
}
