package com.balance.excel.merge.excel;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelCellStyle {
    private String dataFormat;
    private short fillForegroundColor;
    private FillPatternType fillPattern;
    private HorizontalAlignment alignment;
    private BorderStyle borderStyle;

    @Override
    public String toString() {
        return "{" +
                "dataFormat=" + dataFormat +
                ",fillForegroundColor=" + fillForegroundColor +
                ",fillPattern=" + fillPattern +
                ",alignment=" + alignment +
                ",borderStyle=" + borderStyle +
                '}';
    }

    public static ExcelCellStyle fromAnnotation(ExcelStyle annotation) {
        return ExcelCellStyle.builder()
                .dataFormat(annotation.dataFormat())
                .fillForegroundColor(annotation.fillForegroundColor())
                .fillPattern(annotation.fillPattern())
                .alignment(annotation.alignment())
                .borderStyle(annotation.borderStyle())
                .build();
    }

    public static ExcelCellStyle fromDataFormat(String dataFormat) {
        return ExcelCellStyle.builder()
                .dataFormat(dataFormat)
                .build();
    }
}
