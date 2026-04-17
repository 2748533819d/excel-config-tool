package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 范围配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class RangeConfig {

    /**
     * 是否跳过空行
     */
    private Boolean skipEmpty;

    /**
     * 最大读取行数
     */
    private Integer maxRows;

    /**
     * 固定行数（特殊场景使用）
     */
    private Integer rows;

    /**
     * 固定列数（特殊场景使用）
     */
    private Integer cols;

    public RangeConfig() {
    }

    public Boolean getSkipEmpty() {
        return skipEmpty;
    }

    public void setSkipEmpty(Boolean skipEmpty) {
        this.skipEmpty = skipEmpty;
    }

    public Integer getMaxRows() {
        return maxRows;
    }

    public void setMaxRows(Integer maxRows) {
        this.maxRows = maxRows;
    }

    public Integer getRows() {
        return rows;
    }

    public void setRows(Integer rows) {
        this.rows = rows;
    }

    public Integer getCols() {
        return cols;
    }

    public void setCols(Integer cols) {
        this.cols = cols;
    }
}
