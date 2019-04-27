package com.balance.excel.merge.excel;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Getter
@Builder
@AllArgsConstructor
@Slf4j
public class ExcelColumnDefinition {
    private String excelTitle;
    private boolean required;
    private ExcelValueValidator validator;
    private ExeclValueAccessor accessor;
    private ExcelCellStyle cellStyle;

    @Override
    public String toString() {
        return "{" +
                "excelTitle=" + excelTitle +
                ",required=" + required +
                ",accessor=" + accessor +
                ",validator=" + validator +
                ",cellStyle=" + cellStyle +
                '}';
    }

    public static List<ExcelColumnDefinition> parse(Class dataType, Map<String, String> fieldMap) {
        return fieldMap.entrySet().stream()
                .map(entity -> {
                    PropertyDescriptor desc = null;
                    try {
                        desc = new PropertyDescriptor(entity.getValue(), dataType);
                    } catch (IntrospectionException e) {
                        log.error("Can't get field:" + entity.getValue());
                    }
                    return ExcelColumnDefinition.builder()
                            .excelTitle(entity.getKey())
                            .required(false)
                            .accessor(new ExeclValueAccessor(entity.getValue(), desc.getPropertyType(),
                                    desc.getReadMethod(), desc.getWriteMethod()))
                            .build();
                })
                .collect(Collectors.toList());
    }
}
