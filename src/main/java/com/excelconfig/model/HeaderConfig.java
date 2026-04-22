package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 表头匹配配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class HeaderConfig {

    /**
     * 匹配的表头文字
     */
    private String match;

    /**
     * 搜索行范围 [start, end]
     * 不指定则全局搜索
     */
    private int[] inRows;

    public HeaderConfig() {
    }

    public String getMatch() {
        return match;
    }

    public void setMatch(String match) {
        this.match = match;
    }

    public int[] getInRows() {
        return inRows;
    }

    public void setInRows(int[] inRows) {
        this.inRows = inRows;
    }
}
