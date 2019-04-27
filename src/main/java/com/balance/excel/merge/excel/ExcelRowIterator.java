package com.balance.excel.merge.excel;

import com.balance.excel.merge.util.MapUtils;
import com.balance.excel.merge.util.MessageRecorder;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.util.StringUtils;

import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

@Slf4j
public class ExcelRowIterator<T> implements Iterator<ExcelRowWrapper<T>> {
    protected final SheetAgency sheetAgency;
    private ExcelObjectDefinition<T> objectDefinition;
    private int keyFieldIndex = -1;

    protected List<FieldInfo> modelInformation = new ArrayList();

    protected int currentRow;
    protected ExcelRowWrapper<T> currentModel;
    protected MessageRecorder messageRecorder;

    public ExcelRowIterator(Sheet sheet, ExcelObjectDefinition<T> objectDefinition, MessageRecorder messageRecorder) {
        this.sheetAgency = SheetAgency.of(sheet);
        this.objectDefinition = objectDefinition;
        this.messageRecorder = messageRecorder;
    }

    public final void initialize(int rowIndex, int startColumn, int endColumn) {
        currentRow = loadSchema(rowIndex, startColumn, endColumn);
    }
    public int loadSchema(int rowIndex, int startColumn, int endColumn){
        Map<String, ExcelColumnDefinition> fieldMap = MapUtils.listToMap(objectDefinition.getFields(), ExcelColumnDefinition::getExcelTitle);

        for (int index = startColumn; index < endColumn ; ++index) {

            String columnValue = sheetAgency.readCellValue(rowIndex, index, String.class);
            if (!StringUtils.hasText(columnValue)) {
                break;
            }
            ExcelColumnDefinition columnDefinition = fieldMap.get(columnValue.trim());
            if (columnDefinition == null) {
                continue;
            }

            FieldInfo fieldInfo = new FieldInfo(index, columnDefinition);
            // The Key field must at the first.Then it can be checked first when load data
            if (Objects.equals(columnDefinition, objectDefinition.getKeyField())) {
                keyFieldIndex = index;
                modelInformation.add(0, fieldInfo);
            } else {
                modelInformation.add(fieldInfo);
            }

            if (modelInformation.size() == fieldMap.size()) {
                break;
            }
        }
        if (objectDefinition.getKeyField() != null && keyFieldIndex < 0) {
            messageRecorder.appendError(
                    "{0} is not a validate sheet, There is no column with name:{1}.",
                            sheetAgency.getSheet().getSheetName(), objectDefinition.getKeyField().getExcelTitle());
        }

        checkAllNotNullFieldMustExist(objectDefinition, modelInformation);

        return rowIndex + 1;
    }

    private void checkAllNotNullFieldMustExist(ExcelObjectDefinition<T> objectDefinition, List<FieldInfo> modelInformation) {
        Map<String, FieldInfo> fieldInfoMap = MapUtils.listToMap(modelInformation,
                f -> f.getColumnDefinition().getExcelTitle());

        List<String> lostRequiredField = new ArrayList<>();
        for (ExcelColumnDefinition columnDefinition : objectDefinition.getFields()) {
            if (columnDefinition.isRequired() && !fieldInfoMap.containsKey(columnDefinition.getExcelTitle())) {
                lostRequiredField.add(columnDefinition.getExcelTitle());
            }
        }
        if (!lostRequiredField.isEmpty()) {
            messageRecorder.appendError(
                    "{0} is not a validate sheet, These required column are lost:{1}.",
                    sheetAgency.getSheet().getSheetName(), lostRequiredField.toString());
        }
    }

    @Override
    public boolean hasNext() {
        currentModel = loadData(currentRow);
        if (currentModel != null) {
            currentRow++;
            return true;
        }
        return false;
    }

    @Override
    public ExcelRowWrapper<T> next() {
        if (currentModel == null) {
            throw new NoSuchElementException();
        }
        return currentModel;
    }

    public ExcelRowWrapper<T> loadData(int currentRow) {
        if (modelInformation.isEmpty() || keyFieldIndex < 0) {
            return null;
        }
        try {
            T data = objectDefinition.getDataType().newInstance();
            ExcelRowWrapper<T> wrapper = new ExcelRowWrapper<T>(data, sheetAgency.getSheet().getSheetName(), currentRow + 1);
            for (FieldInfo fieldInfo : modelInformation) {
                Object value = sheetAgency.readCellValue(currentRow, fieldInfo.columnIndex,
                        fieldInfo.getColumnDefinition().getAccessor().getDataType());

                boolean isEmptyValue = ExcelValueValidator.isEmptyValue(value);
                if (fieldInfo.columnIndex == keyFieldIndex && isEmptyValue) {
                    return null;
                }

                if (fieldInfo.columnDefinition.getValidator() != null) {
                    fieldInfo.columnDefinition.getValidator().validate(value, fieldInfo.columnDefinition, wrapper, messageRecorder);
                }

                if (!isEmptyValue) {
                    try {
                        fieldInfo.columnDefinition.getAccessor().getSetMethod().invoke(data, value);
                    } catch (IllegalAccessException | InvocationTargetException e) {
                        log.error("Failed to write field:" + fieldInfo.columnDefinition.getAccessor().getFieldName(), e);
                    }
                }
            }

            objectDefinition.runPostAction(wrapper, messageRecorder);
            return wrapper;
        } catch (InstantiationException | IllegalAccessException e) {
            log.error("Failed to create instance " + objectDefinition.getDataType(), e);
        }
        return null;
    }

    public Stream<ExcelRowWrapper<T>> stream() {
        Iterable<ExcelRowWrapper<T>> iterable = () -> this;
        return StreamSupport.stream(iterable.spliterator(), false);
    }

    public static <T> ExcelRowIterator<T> from(Sheet sheet, Class<T> dataType) {
        return from(sheet, ExcelObjectDefinition.from(dataType), new MessageRecorder());
    }

    public static <T> ExcelRowIterator<T> from(Sheet sheet, Class<T> dataType, MessageRecorder messageRecorder) {
        return from(sheet, ExcelObjectDefinition.from(dataType), messageRecorder);
    }

    public static <T> ExcelRowIterator<T> from(Sheet sheet, ExcelObjectDefinition<T> objectDefinition, MessageRecorder messageRecorder) {
        ExcelRowIterator<T> iterator = new ExcelRowIterator<>(sheet, objectDefinition, messageRecorder);
        Row row = sheet.getRow(0);
        if (row != null) {
            iterator.initialize(0, 0, row.getLastCellNum());
        }
        return iterator;
    }

    @Getter
    @AllArgsConstructor
    private static class FieldInfo {
        private int columnIndex;
        private ExcelColumnDefinition columnDefinition;
    }
}
