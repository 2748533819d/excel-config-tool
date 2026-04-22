package com.excelconfig.spi;

import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.PositionConfig;

import java.io.InputStream;

/**
 * 提取上下文
 */
public class ExtractContext {

    /**
     * 当前配置
     */
    private final ExtractConfig config;

    /**
     * 起始行
     */
    private final int startRow;

    /**
     * 起始列
     */
    private final int startColumn;

    /**
     * 结束行
     */
    private Integer endRow;

    /**
     * 结束列
     */
    private Integer endColumn;

    /**
     * Sheet 索引
     */
    private int sheetIndex = 0;

    /**
     * 输入流（用于 SAX 流式读取）
     */
    private InputStream inputStream;

    public ExtractContext(ExtractConfig config, int startRow, int startColumn) {
        this.config = config;
        this.startRow = startRow;
        this.startColumn = startColumn;
    }

    public ExtractContext(ExtractConfig config, int startRow, int startColumn, int sheetIndex) {
        this.config = config;
        this.startRow = startRow;
        this.startColumn = startColumn;
        this.sheetIndex = sheetIndex;
    }

    public ExtractConfig getConfig() {
        return config;
    }

    public int getStartRow() {
        return startRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public Integer getEndRow() {
        return endRow;
    }

    public void setEndRow(Integer endRow) {
        this.endRow = endRow;
    }

    public Integer getEndColumn() {
        return endColumn;
    }

    public void setEndColumn(Integer endColumn) {
        this.endColumn = endColumn;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }
}
