package com.excelconfig.spi;

/**
 * 提取模式枚举
 */
public enum ExtractMode {
    // 基础模式
    SINGLE,         // 单个单元格
    DOWN,           // 向下提取
    RIGHT,          // 向右提取
    BLOCK,          // 区域矩阵
    UNTIL_EMPTY,    // 提取到空行

    // 扩展模式
    KEY_VALUE,      // 键值对
    TABLE,          // 表格
    CROSS_TAB,      // 交叉表
    GROUPED,        // 分组
    HIERARCHY,      // 层级
    MERGED_CELLS,   // 合并单元格
    MULTI_SHEET,    // 多工作表
    PIVOT,          // 透视表
    FORMULA,        // 公式
    CONDITIONAL;    // 条件

    public static ExtractMode fromString(String value) {
        if (value == null) {
            return null;
        }
        try {
            return ExtractMode.valueOf(value.toUpperCase());
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("未知的提取模式：" + value);
        }
    }
}
