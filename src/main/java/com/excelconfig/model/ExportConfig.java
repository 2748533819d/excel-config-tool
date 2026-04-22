package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 导出配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class ExportConfig {

    /**
     * 字段名
     */
    private String key;

    /**
     * 表头匹配配置
     */
    private HeaderConfig header;

    /**
     * 固定位置配置
     */
    private PositionConfig position;

    /**
     * 导出模式
     */
    private String mode;

    /**
     * 列配置（用于 FILL_TABLE 等模式）
     */
    private java.util.List<ColumnConfig> columns;

    /**
     * 样式配置
     */
    private StyleConfig style;

    /**
     * 表头样式配置
     */
    private StyleConfig headerStyle;

    /**
     * 最大填充行数
     */
    private Integer maxRows;

    /**
     * 是否隔行换色
     */
    private Boolean alternateRows;

    /**
     * 是否自动列宽
     */
    private Boolean autoWidth;

    /**
     * 合并单元格配置
     */
    private MergeConfig merge;

    public ExportConfig() {
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

    public java.util.List<ColumnConfig> getColumns() {
        return columns;
    }

    public void setColumns(java.util.List<ColumnConfig> columns) {
        this.columns = columns;
    }

    public StyleConfig getStyle() {
        return style;
    }

    public void setStyle(StyleConfig style) {
        this.style = style;
    }

    public StyleConfig getHeaderStyle() {
        return headerStyle;
    }

    public void setHeaderStyle(StyleConfig headerStyle) {
        this.headerStyle = headerStyle;
    }

    public Integer getMaxRows() {
        return maxRows;
    }

    public void setMaxRows(Integer maxRows) {
        this.maxRows = maxRows;
    }

    public Boolean getAlternateRows() {
        return alternateRows;
    }

    public void setAlternateRows(Boolean alternateRows) {
        this.alternateRows = alternateRows;
    }

    public Boolean getAutoWidth() {
        return autoWidth;
    }

    public void setAutoWidth(Boolean autoWidth) {
        this.autoWidth = autoWidth;
    }

    public MergeConfig getMerge() {
        return merge;
    }

    public void setMerge(MergeConfig merge) {
        this.merge = merge;
    }
}
