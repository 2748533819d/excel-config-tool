package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 列配置（用于 FILL_TABLE 等模式）
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class ColumnConfig {

    /**
     * 字段名
     */
    private String key;

    /**
     * 表头文字
     */
    private String header;

    /**
     * 列宽
     */
    private Integer width;

    /**
     * 格式
     */
    private String format;

    /**
     * 样式配置
     */
    private StyleConfig style;

    public ColumnConfig() {
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public String getHeader() {
        return header;
    }

    public void setHeader(String header) {
        this.header = header;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public StyleConfig getStyle() {
        return style;
    }

    public void setStyle(StyleConfig style) {
        this.style = style;
    }
}
