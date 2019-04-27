package com.balance.excel.merge.util.convert.primitive;

import java.text.MessageFormat;
import java.util.HashMap;
import java.util.Map;


public class PrimitiveConverter {
    private static final PrimitiveConverter instance = new PrimitiveConverter();

    public static PrimitiveConverter getInstance() {
        return instance;
    }

    private Map<Class, IPrimitiveConverter> converterMap = new HashMap<>();
    public <T> T convert(Object source, Class<T> targetType) {
        IPrimitiveConverter targetConverter = findConverter(targetType);
        if (targetConverter == null) {
            throw new UnsupportedOperationException(
                    MessageFormat.format("Can''t convert {0} to {1} for no convert define for {1}.",
                            source, targetType));
        }

        String strSource = null;
        if (source != null) {
            IPrimitiveConverter sourceConverter = findConverter(source.getClass());
            if (sourceConverter == null) {
                throw new UnsupportedOperationException(
                        MessageFormat.format("Can''t convert {0} to {1} for no convert define for {2}.",
                                source, targetType, source.getClass()));
            }
            strSource = sourceConverter.convertToString(source);
        }
        return (T) targetConverter.convertFromString(strSource);
    }

    public void register(IPrimitiveConverter converter) {
        converterMap.put(converter.dataType(), converter);
        Class primitiveType = converter.primitiveType();
        if (primitiveType != null) {
            converterMap.put(primitiveType, converter);
        }
    }

    public IPrimitiveConverter findConverter(Class dataType) {
        IPrimitiveConverter converter = converterMap.get(dataType);
        if (converter == null && dataType.isEnum()) {
            converter = new EnumConverter(dataType);
            register(converter);
        }
        return converter;
    }

    public boolean isPrimitive(Class dataType) {
        return findConverter(dataType) != null;
    }
}
