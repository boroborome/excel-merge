package com.balance.excel.merge;

import com.balance.excel.merge.excel.SheetAgency;
import com.balance.excel.merge.model.ExcelSheetData;
import com.balance.excel.merge.model.Pair;
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
                .filter(p -> Objects.nonNull(p.getValue()))
                .flatMap(nameWorkBookPair -> listAllSheets(nameWorkBookPair))
                .map(sheetPair -> ExcelSheetData.parseFromSheet(sheetPair.getValue(), sheetPair.getKey()))
                .filter(excelData -> !CollectionUtils.isEmpty(excelData.getRows()))
                .reduce(new ExcelSheetData(),
                        (a, b) -> a.merge(b));

        saveToExcel(excelSheetData, summaryFile);
    }

    private void saveToExcel(ExcelSheetData excelSheetData, File summaryFile) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        SheetAgency sheetAgency = SheetAgency.of(workbook, "summary");
        excelSheetData.getTitles().add(ExcelSheetData.EXCEL_LOCATION);
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

    private Stream<Pair<String, Sheet>> listAllSheets(Pair<String, Workbook> nameWorkbookPair) {
        Workbook workbook = nameWorkbookPair.getValue();
        int count = workbook.getNumberOfSheets();
        List<Pair<String, Sheet>> sheets = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            sheets.add(new Pair<>(
                    nameWorkbookPair.getKey(),
                    sheet));
        }
        return sheets.stream();
    }

    private Pair<String, Workbook> openExcel(File file) {
        Workbook workbook = openWithHSSF(file);
        if (workbook == null) {
            workbook = openWithXssf(file);
        }
        return new Pair<>(file.getName(), workbook);

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
