package com.excelconfig.spi;

import com.excelconfig.model.ExportConfig;
import java.util.Map;

/**
 * 填充上下文
 */
public class FillContext {

    /**
     * 当前配置
     */
    private final ExportConfig config;

    /**
     * 数据
     */
    private final Map<String, Object> data;

    /**
     * 起始行
     */
    private final int startRow;

    /**
     * 起始列
     */
    private final int startColumn;

    public FillContext(ExportConfig config, Map<String, Object> data, int startRow, int startColumn) {
        this.config = config;
        this.data = data;
        this.startRow = startRow;
        this.startColumn = startColumn;
    }

    public ExportConfig getConfig() {
        return config;
    }

    public Map<String, Object> getData() {
        return data;
    }

    public int getStartRow() {
        return startRow;
    }

    public int getStartColumn() {
        return startColumn;
    }
}
