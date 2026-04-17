package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 解析器配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class ParserConfig {

    /**
     * 解析器类型：string, number, date, boolean
     */
    private String type;

    /**
     * 格式（用于 date 和 number 类型）
     */
    private String format;

    /**
     * 小数位数（用于 number 类型）
     */
    private Integer scale;

    public ParserConfig() {
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public Integer getScale() {
        return scale;
    }

    public void setScale(Integer scale) {
        this.scale = scale;
    }
}
