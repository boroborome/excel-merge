package com.balance.excel.merge;

import com.balance.excel.merge.excel.SheetAgency;
import com.balance.excel.merge.model.ExcelSheetData;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

@Slf4j
public class ExcelMerger {
    private final List<File> inputFiles;
    private final File summaryFile;

    public ExcelMerger(File[] inputFiles, File summaryFile) {
        this.inputFiles = Arrays.asList(inputFiles);
        this.summaryFile = summaryFile;
    }

    public void merge() {
        ExcelSheetData excelSheetData = inputFiles.stream()
                .map(file -> openExcel(file))
                .filter(Objects::nonNull)
                .flatMap(workBook -> listAllSheets(workBook))
                .map(sheet -> ExcelSheetData.parseFromSheet(sheet))
                .filter(excelData -> !CollectionUtils.isEmpty(excelData.getRows()))
                .reduce(new ExcelSheetData(),
                        (a, b) -> a.merge(b));

        saveToExcel(excelSheetData, summaryFile);
    }

    private void saveToExcel(ExcelSheetData excelSheetData, File summaryFile) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        SheetAgency sheetAgency = SheetAgency.of(workbook, "summary");
        sheetAgency.write(excelSheetData.getTitles().toArray())
                .newLine();

        List<String> values = new ArrayList<>();
        for (Map<String, String> row : excelSheetData.getRows()) {
            values.clear();
            for (String title : excelSheetData.getTitles()) {
                values.add(row.get(title));
            }

            sheetAgency.write(values.toArray())
                    .newLine();
        }

        try {
            workbook.write(new FileOutputStream(summaryFile));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Stream<Sheet> listAllSheets(Workbook workBook) {
        int count = workBook.getNumberOfSheets();
        List<Sheet> sheets = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            sheets.add(workBook.getSheetAt(i));
        }
        return sheets.stream();
    }

    private Workbook openExcel(File file) {
        Workbook workbook = openWithHSSF(file);
        if (workbook == null) {
            workbook = openWithXssf(file);
        }
        return workbook;

    }

    private Workbook openWithXssf(File file) {
        try {
            return new XSSFWorkbook(file);
        } catch (Exception e) {
            log.error("Failed to open file:{} with format XSSF.", file.getAbsolutePath());
        }
        return null;
    }

    private Workbook openWithHSSF(File file) {
        try {
            return new HSSFWorkbook(new FileInputStream(file));
        } catch (Exception e) {
            log.error("Failed to open file:{} with format HSSF.", file.getAbsolutePath());
        }
        return null;
    }
}
