package com.balance.excel.merge.excel;

import com.balance.excel.merge.util.MessageRecorder;
import lombok.AllArgsConstructor;
import lombok.Builder;
import org.springframework.util.StringUtils;

@AllArgsConstructor
@Builder
public class ExcelValueValidator {
    private boolean notEmpty;

    public <T> void validate(Object value, ExcelColumnDefinition columnDefinition, ExcelRowWrapper<T> wrapper, MessageRecorder messageRecorder) {
        if (notEmpty && isEmptyValue(value)) {
            messageRecorder.appendError(
                    "Field:{0} is required at line {1} in sheet {2}.",
                    columnDefinition.getExcelTitle(),
                    wrapper.getRowIndex(),
                    wrapper.getSheetName());
        }
    }

    @Override
    public String toString() {
        return "{" +
                "notEmpty=" + notEmpty +
                '}';
    }

    public static boolean isEmptyValue(Object value) {
        if (value == null) {
            return true;
        }
        if (value instanceof String) {
            String s = (String) value;
            return StringUtils.isEmpty(s);
        }
        return false;
    }

    public static ExcelValueValidator from(ExcelValueValidation valueValidation) {
        if (valueValidation == null) {
            return null;
        }

        return new ExcelValueValidator(valueValidation.notEmpty());
    }
}
