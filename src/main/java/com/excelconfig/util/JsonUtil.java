package com.excelconfig.util;

import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;

import java.util.Map;

/**
 * JSON 工具类 - 基于 Jackson ObjectMapper
 */
public class JsonUtil {

    private static final ObjectMapper objectMapper;

    static {
        objectMapper = new ObjectMapper();
        // 忽略不存在的字段
        objectMapper.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);
        // 忽略未知枚举值
        objectMapper.configure(DeserializationFeature.READ_UNKNOWN_ENUM_VALUES_AS_NULL, true);
        // 支持 Java 8 日期时间类型
        objectMapper.registerModule(new JavaTimeModule());
        // 使用 ISO-8601 格式
        objectMapper.configure(DeserializationFeature.ADJUST_DATES_TO_CONTEXT_TIME_ZONE, true);
    }

    private JsonUtil() {
        // 工具类，禁止实例化
    }

    /**
     * 获取 ObjectMapper 实例
     *
     * @return ObjectMapper
     */
    public static ObjectMapper getObjectMapper() {
        return objectMapper;
    }

    /**
     * 将对象转换为 JSON 字符串
     *
     * @param obj 要转换的对象
     * @return JSON 字符串
     */
    public static String toJson(Object obj) {
        try {
            return objectMapper.writeValueAsString(obj);
        } catch (Exception e) {
            throw new RuntimeException("转换为 JSON 失败：" + e.getMessage(), e);
        }
    }

    /**
     * 将 JSON 字符串转换为对象
     *
     * @param json JSON 字符串
     * @param clazz 目标类型
     * @param <T> 类型参数
     * @return 转换后的对象
     */
    public static <T> T fromJson(String json, Class<T> clazz) {
        try {
            return objectMapper.readValue(json, clazz);
        } catch (Exception e) {
            throw new RuntimeException("解析 JSON 失败：" + e.getMessage(), e);
        }
    }

    /**
     * 将 Map 转换为对象（忽略 Map 中目标类不存在的字段）
     *
     * @param map 源 Map
     * @param clazz 目标类型
     * @param <T> 类型参数
     * @return 转换后的对象
     */
    public static <T> T convertToObject(Map<String, Object> map, Class<T> clazz) {
        return objectMapper.convertValue(map, clazz);
    }

    /**
     * 将对象转换为 Map
     *
     * @param obj 源对象
     * @return Map
     */
    public static Map<String, Object> convertToMap(Object obj) {
        return objectMapper.convertValue(obj, Map.class);
    }
}
