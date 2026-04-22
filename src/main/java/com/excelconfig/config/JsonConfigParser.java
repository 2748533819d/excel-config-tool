package com.excelconfig.config;

import com.excelconfig.model.ExcelConfig;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.DeserializationFeature;

import java.io.InputStream;
import java.io.IOException;

/**
 * JSON 配置解析器
 *
 * 将 JSON 配置文件解析为 ExcelConfig 对象
 */
public class JsonConfigParser {

    private final ObjectMapper objectMapper;

    public JsonConfigParser() {
        this.objectMapper = new ObjectMapper();
        // 配置 ObjectMapper
        this.objectMapper.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);
    }

    /**
     * 从 InputStream 解析配置
     *
     * @param inputStream JSON 配置输入流
     * @return ExcelConfig 配置对象
     * @throws IOException 读取失败
     */
    public ExcelConfig parse(InputStream inputStream) throws IOException {
        return objectMapper.readValue(inputStream, ExcelConfig.class);
    }

    /**
     * 从 JSON 字符串解析配置
     *
     * @param json JSON 配置字符串
     * @return ExcelConfig 配置对象
     * @throws IOException 解析失败
     */
    public ExcelConfig parse(String json) throws IOException {
        return objectMapper.readValue(json, ExcelConfig.class);
    }
}
