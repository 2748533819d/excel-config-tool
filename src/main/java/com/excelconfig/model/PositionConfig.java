package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 位置配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class PositionConfig {

    /**
     * 单元格引用，如 "A1", "B2"
     */
    private String cellRef;

    public PositionConfig() {
    }

    public String getCellRef() {
        return cellRef;
    }

    public void setCellRef(String cellRef) {
        this.cellRef = cellRef;
    }
}
