package com.balance.excel.merge.util;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.springframework.util.CollectionUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Supplier;

public class MapUtils {
    public static <T_VALUE, T_KEY> Map<T_VALUE, T_KEY> swapToValueKey(Map<T_KEY, T_VALUE> map) {
        Map<T_VALUE, T_KEY> resultMap = new HashMap<>();

        map.forEach((k, v) -> resultMap.put(v, k));

        return resultMap;
    }

    public static <K,V> Map<K, V> from(List<V> list, Function<V, K> keyGenerator) {
        Map<K, V> map = new HashMap<>();
        if (CollectionUtils.isEmpty(list)) {
            return map;
        }
        for (V v : list) {
            K k = keyGenerator.apply(v);
            map.put(k, v);
        }

        return map;
    }

    public static <K, V> V safeRead(K key, Supplier<V> supplier, Map<K, V> map) {
        V value = map.get(key);
        if (value == null) {
            value = supplier.get();
            map.put(key, value);
        }
        return value;
    }
    public static <K, V> V safeRead(Map<K, V> map, K key, Supplier<V> supplier) {
        V value = map.get(key);
        if (value == null) {
            value = supplier.get();
            map.put(key, value);
        }
        return value;
    }
    public static <V> V safeRead(Map<?, ?> map, Supplier<V> supplierLast,Object... otherKeys) {
        List<MapKey> mapKeys = new ArrayList<>();
        if (otherKeys == null || otherKeys.length <= 0) {
            throw new IllegalArgumentException("At lest has one key.");
        } else {
            for (int i = 0; i < otherKeys.length - 1; i++) {
                Object key = otherKeys[i];
                mapKeys.add(MapKey.of(key, () -> new HashMap()));
            }
            mapKeys.add(MapKey.of(otherKeys[otherKeys.length - 1], supplierLast));
        }
        return (V) safeRead(map, mapKeys.toArray(new MapKey[otherKeys.length]));
    }

    public static <K, V> V safeRead(Map<K, V> map, MapKey ... mapKeys) {
        Object v = map;
        for (MapKey mapKey : mapKeys) {
            Map m = (Map) v;
            v = safeRead(m, mapKey.key, mapKey.supplier);
        }
        return (V) v;
    }

    @Getter
    @AllArgsConstructor
    public static class MapKey <K, V> {
        private K key;
        private Supplier<V> supplier;
        public static <K, V> MapKey<K, V> of(K key, Supplier<V> supplier) {
            return new MapKey<>(key, supplier);
        }
    }

    public static <K,T> Map<K, T> listToMap(List<T> list, Function<T, K> keyGenerator) {
        return listToMap(list, keyGenerator, e -> e);
    }

    public static <K, V, T> Map<K, V> listToMap(List<T> list, Function<T, K> keyGenerator, Function<T, V> valueGenerator) {
        Map<K, V> map = new HashMap<>();
        if (list != null) {
            for (T item : list) {
                map.put(keyGenerator.apply(item), valueGenerator.apply(item));
            }
        }
        return map;
    }
}
