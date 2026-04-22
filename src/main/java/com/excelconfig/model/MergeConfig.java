package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 合并单元格配置
 *
 * 支持两种合并模式：
 * 1. 固定区域合并：按 rowSpan/colSpan 合并指定区域
 * 2. 智能合并：根据数据值自动合并相同值的单元格
 *
 * 示例 1 - 智能合并（按数据值）：
 * <pre>
 * {
 *   "key": "department",
 *   "mode": "FILL_DOWN",
 *   "merge": {
 *     "enabled": true
 *   }
 * }
 * </pre>
 * 数据：["技术部", "技术部", "技术部", "销售部", "销售部"]
 * 效果：A1:A3 合并为"技术部"，A4:A5 合并为"销售部"
 *
 * 示例 2 - 固定区域合并：
 * <pre>
 * {
 *   "key": "title",
 *   "mode": "FILL_DOWN",
 *   "merge": {
 *     "rowSpan": 2,
 *     "colSpan": 2
 *   }
 * }
 * </pre>
 * 效果：每个数据都合并 A1:B2 区域
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class MergeConfig {

    /**
     * 是否启用合并
     * true = 按数据值智能合并
     */
    private Boolean enabled;

    /**
     * 跨越的行数（固定区域合并模式）
     */
    private Integer rowSpan;

    /**
     * 跨越的列数（固定区域合并模式）
     */
    private Integer colSpan;

    /**
     * 合并起始行偏移（相对于数据起始行）
     */
    private Integer startRowOffset;

    /**
     * 合并起始列偏移（相对于数据起始列）
     */
    private Integer startColOffset;

    /**
     * 最小合并数量（智能合并模式）
     * 只有连续相同值的数量 >= 此值才合并
     * 默认值为 2
     */
    private Integer minSpan;

    /**
     * 最大合并数量（智能合并模式）
     * 限制单次合并的最大行数
     */
    private Integer maxSpan;

    public MergeConfig() {
    }

    public MergeConfig(boolean enabled) {
        this.enabled = enabled;
    }

    public MergeConfig(int rowSpan, int colSpan) {
        this.rowSpan = rowSpan;
        this.colSpan = colSpan;
    }

    public Boolean getEnabled() {
        return enabled;
    }

    public void setEnabled(Boolean enabled) {
        this.enabled = enabled;
    }

    public Integer getRowSpan() {
        return rowSpan;
    }

    public void setRowSpan(Integer rowSpan) {
        this.rowSpan = rowSpan;
    }

    public Integer getColSpan() {
        return colSpan;
    }

    public void setColSpan(Integer colSpan) {
        this.colSpan = colSpan;
    }

    public Integer getStartRowOffset() {
        return startRowOffset;
    }

    public void setStartRowOffset(Integer startRowOffset) {
        this.startRowOffset = startRowOffset;
    }

    public Integer getStartColOffset() {
        return startColOffset;
    }

    public void setStartColOffset(Integer startColOffset) {
        this.startColOffset = startColOffset;
    }

    public Integer getMinSpan() {
        return minSpan;
    }

    public void setMinSpan(Integer minSpan) {
        this.minSpan = minSpan;
    }

    public Integer getMaxSpan() {
        return maxSpan;
    }

    public void setMaxSpan(Integer maxSpan) {
        this.maxSpan = maxSpan;
    }

    /**
     * 是否是智能合并模式（按数据值合并）
     */
    public boolean isSmartMerge() {
        return Boolean.TRUE.equals(enabled);
    }

    /**
     * 是否是固定区域合并模式
     */
    public boolean isFixedMerge() {
        return rowSpan != null || colSpan != null;
    }
}
