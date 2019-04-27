package com.balance.excel.merge.util.convert.primitive;

import org.springframework.util.StringUtils;

import java.util.Date;

public class DateConverter implements IPrimitiveConverter<Date> {
    @Override
    public Class<Date> dataType() {
        return Date.class;
    }
    @Override
    public Class primitiveType() {
        return Date.class;
    }

    @Override
    public String convertToString(Date value) {
        return value == null ? null : String.valueOf(value.getTime());
    }

    @Override
    public Date convertFromString(String str) {
        return StringUtils.isEmpty(str) ? null : new Date(Long.parseLong(str));
    }
}
