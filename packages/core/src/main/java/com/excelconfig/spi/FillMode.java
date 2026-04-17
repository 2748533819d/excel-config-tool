package com.excelconfig.spi;

/**
 * 导出模式枚举
 */
public enum FillMode {
    // 基础模式
    FILL_CELL,      // 填充单个单元格
    FILL_DOWN,      // 向下填充
    FILL_RIGHT,     // 向右填充
    FILL_BLOCK,     // 填充区域

    // 表格模式
    FILL_TABLE,     // 填充表格（带表头）
    APPEND_ROWS,    // 追加行
    APPEND_COLS,    // 追加列

    // 高级模式
    REPLACE_AREA,   // 替换区域
    FILL_TEMPLATE,  // 模板填充
    MULTI_SHEET_FILL; // 多工作表填充

    public static FillMode fromString(String value) {
        if (value == null) {
            return null;
        }
        try {
            return FillMode.valueOf(value.toUpperCase());
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("未知的导出模式：" + value);
        }
    }
}
