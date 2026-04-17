package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 提取配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class ExtractConfig {

    /**
     * 字段名（映射到 JSON 结果的 key）
     */
    private String key;

    /**
     * 表头匹配配置
     */
    private HeaderConfig header;

    /**
     * 固定位置配置（备选方式）
     */
    private PositionConfig position;

    /**
     * 提取模式
     */
    private String mode;

    /**
     * 范围配置
     */
    private RangeConfig range;

    /**
     * 解析器配置
     */
    private ParserConfig parser;

    public ExtractConfig() {
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public HeaderConfig getHeader() {
        return header;
    }

    public void setHeader(HeaderConfig header) {
        this.header = header;
    }

    public PositionConfig getPosition() {
        return position;
    }

    public void setPosition(PositionConfig position) {
        this.position = position;
    }

    public String getMode() {
        return mode;
    }

    public void setMode(String mode) {
        this.mode = mode;
    }

    public RangeConfig getRange() {
        return range;
    }

    public void setRange(RangeConfig range) {
        this.range = range;
    }

    public ParserConfig getParser() {
        return parser;
    }

    public void setParser(ParserConfig parser) {
        this.parser = parser;
    }
}
