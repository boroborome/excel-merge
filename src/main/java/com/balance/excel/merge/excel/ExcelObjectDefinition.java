package com.balance.excel.merge.excel;

import com.balance.excel.merge.util.MessageRecorder;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Getter
@Setter
@AllArgsConstructor
@Builder
@Slf4j
public class ExcelObjectDefinition<T> {
    private Class<T> dataType;
    private List<ExcelColumnDefinition> fields;
    private ExeclValueAccessor otherValueAccessor;
    private Method postAction;
    private ExcelColumnDefinition keyField;

    public void runPostAction(ExcelRowWrapper<T> wrapper, MessageRecorder messageRecorder) {
        if (postAction == null) {
            return;
        }

        List<Object> params = new ArrayList<>();
        for (Class paramType : postAction.getParameterTypes()) {
            if (paramType == ExcelRowWrapper.class) {
                params.add(wrapper);
            } else if (paramType == MessageRecorder.class) {
                params.add(messageRecorder);
            } else {
                params.add(null);
            }
        }
        try {
            postAction.invoke(wrapper.getData(), params.toArray());
        } catch (IllegalAccessException | InvocationTargetException e) {
            messageRecorder.appendError("Failed to run post action of {0} for object {1}, Error is :{2}",
                    dataType, wrapper.getData(), e.getMessage());
            log.error(MessageFormat.format("Failed to run post action of {0} for object {1}, Error is :{2}",
                    dataType, wrapper.getData(), e.getMessage(), e));
        }
    }

    public static <T> ExcelObjectDefinition<T> from(Class<T> dataType) {
        ExeclValueAccessor otherValueAccessor = null;
        List<ExcelColumnDefinition> fields = new ArrayList<>();
        ExcelColumnDefinition keyField = null;

        Method[] methods = dataType.getDeclaredMethods();
        for (Field field : dataType.getDeclaredFields()) {

            ExcelOtherColumn[] otherColumnAnnotations = field.getAnnotationsByType(ExcelOtherColumn.class);
            if (otherColumnAnnotations.length > 0) {
                if (otherValueAccessor != null) {
                    throw new UnsupportedOperationException("Unsupported multiple ExcelOtherColumn configured in one class:" + dataType);
                }
                otherValueAccessor = ExeclValueAccessor.from(field);
                continue;
            }

            ExcelColumn[] columnAnnotations = field.getAnnotationsByType(ExcelColumn.class);
            if (columnAnnotations.length == 0) {
                continue;
            }
            ExcelColumn excelColumn = columnAnnotations[0];
            ExeclValueAccessor accessor = getExeclValueAccessor(methods, field, excelColumn);
            ExcelValueValidation valueValidation = getExcelValueValidation(field);
            ExcelCellStyle cellStyle = getExcelStyle(field);

            ExcelColumnDefinition columnDef = ExcelColumnDefinition.builder()
                    .excelTitle(excelColumn.value())
                    .required(excelColumn.required())
                    .accessor(accessor)
                    .cellStyle(cellStyle)
                    .validator(ExcelValueValidator.from(valueValidation))
                    .build();
            fields.add(columnDef);

            ExcelKeyField[] keyFieldAnnotations = field.getAnnotationsByType(ExcelKeyField.class);
            if (keyFieldAnnotations.length > 0) {
                if (keyField != null) {
                    throw new UnsupportedOperationException("Unsupported multiple ExcelKeyField configured in one class:" + dataType);
                }
                keyField = columnDef;
            }
        }

        Method postAction = getPostAction(dataType);
        return ExcelObjectDefinition.<T>builder()
                .dataType(dataType)
                .fields(fields)
                .keyField(keyField)
                .otherValueAccessor(otherValueAccessor)
                .postAction(postAction)
                .build();
    }

    private static ExcelCellStyle getExcelStyle(Field field) {
        ExcelStyle[] styleAnnotations = field.getAnnotationsByType(ExcelStyle.class);
        return styleAnnotations.length > 0 ? ExcelCellStyle.fromAnnotation(styleAnnotations[0]) : null;
    }

    private static ExcelValueValidation getExcelValueValidation(Field field) {
        ExcelValueValidation[] valueValidations = field.getAnnotationsByType(ExcelValueValidation.class);
        return valueValidations.length > 0 ? valueValidations[0] : null;
    }

    private static ExeclValueAccessor getExeclValueAccessor(Method[] methods, Field field, ExcelColumn excelColumn) {
        ExeclValueAccessor accessor = null;
        if (StringUtils.isEmpty(excelColumn.getter()) || StringUtils.isEmpty(excelColumn.setter())) {
            accessor = ExeclValueAccessor.from(field);
        }
        Method getter = StringUtils.isEmpty(excelColumn.getter()) ? accessor.getGetMethod() : findMethod(excelColumn.getter(), methods);
        Method setter = StringUtils.isEmpty(excelColumn.setter()) ? accessor.getSetMethod() : findMethod(excelColumn.setter(), methods);
        if (accessor == null || accessor.getGetMethod() != getter || accessor.getSetMethod() != setter) {
            accessor = new ExeclValueAccessor(field.getName(), getter.getReturnType(), getter, setter);
        }
        return accessor;
    }

    private static Method findMethod(String methodName, Method[] methods) {
        for (Method method : methods) {
            if (Objects.equals(method.getName(), methodName)) {
                return method;
            }
        }
        throw new IllegalArgumentException("Can't find method:" + methodName);
    }

    private static Method getPostAction(Class dataType) {
        Method postAction = null;
        Method[] methods = dataType.getDeclaredMethods();
        for (Method method : methods) {
            ExcelPostAction[] postActionAnnotations = method.getAnnotationsByType(ExcelPostAction.class);
            if (postActionAnnotations.length > 0) {
                if (postAction != null) {
                    throw new UnsupportedOperationException("Unsupported multiple ExcelKeyField configured in one class:" + dataType);
                }
                postAction = method;
            }
        }
        return postAction;
    }

    @Override
    public String toString() {
        return "{" +
                "dataType=" + dataType +
                ",fields=" + fields +
                ",otherValueAccessor=" + otherValueAccessor +
                ",postAction=" + postAction +
                ",keyField=" + keyField +
                '}';
    }
}
