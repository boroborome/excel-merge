package com.balance.excel.merge.util.convert.primitive;

import org.springframework.util.StringUtils;

import java.sql.Timestamp;
import java.time.Instant;

public class TimestampConverter implements IPrimitiveConverter<Timestamp> {
    @Override
    public Class<Timestamp> dataType() {
        return Timestamp.class;
    }
    @Override
    public Class primitiveType() {
        return Timestamp.class;
    }

    @Override
    public String convertToString(Timestamp value) {
        return value == null ? null : String.valueOf(value.getTime());
    }

    @Override
    public Timestamp convertFromString(String str) {
        return StringUtils.isEmpty(str) ? null : Timestamp.from(Instant.ofEpochMilli(Long.parseLong(str)));
    }
}
