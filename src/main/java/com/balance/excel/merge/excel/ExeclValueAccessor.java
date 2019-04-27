package com.balance.excel.merge.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

@Getter
@AllArgsConstructor
@Slf4j
public class ExeclValueAccessor {
    private String fieldName;
    private Class dataType;
    private Method getMethod;
    private Method setMethod;

    @Override
    public String toString() {
        return "{" +
                "fieldName=" + fieldName +
                ",dataType=" + dataType +
                ",getMethod=" + (getMethod == null ? null : getMethod.getName()) +
                ",setMethod=" + (setMethod == null ? null : setMethod.getName()) +
                '}';
    }

    public static ExeclValueAccessor from(Field field) {
        try {
            PropertyDescriptor desc = new PropertyDescriptor(field.getName(), field.getDeclaringClass());
            return new ExeclValueAccessor(field.getName(), field.getType(),
                            desc.getReadMethod(), desc.getWriteMethod());
        } catch (IntrospectionException e) {
            throw new RuntimeException("No Property define for field:" + field.toString(), e);
        }
    }

    public Object getValue(Object data) {
        if (data == null) {
            return null;
        }
        try {
            return getMethod.invoke(data);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException("Failed to read field:" + fieldName + " from:" + data, e);
        }
    }
}
