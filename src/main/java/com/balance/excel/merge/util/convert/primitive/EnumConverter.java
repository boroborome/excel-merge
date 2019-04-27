package com.balance.excel.merge.util.convert.primitive;

import java.util.HashMap;
import java.util.Map;

public class EnumConverter<T extends Enum> implements IPrimitiveConverter<T> {
    private Class<T> dataType;
    private Map<String, T> nameToValueMap = new HashMap<>();
    private Map<T, String> valueToNameMap = new HashMap<>();

    public EnumConverter(Class<T> dataType) {
        this.dataType = dataType;
        for (T v : dataType.getEnumConstants()) {
            nameToValueMap.put(v.name(), v);
            valueToNameMap.put(v, v.name());
        }
    }

    @Override
    public Class<T> dataType() {
        return dataType;
    }

    @Override
    public Class primitiveType() {
        return null;
    }

    @Override
    public String convertToString(T value) {
        return valueToNameMap.get(value);
    }

    @Override
    public T convertFromString(String name) {
        return nameToValueMap.get(name);
    }
}
