package com.excelconfig.export;

import com.excelconfig.model.ColumnConfig;
import com.excelconfig.model.ExportConfig;
import com.excelconfig.model.MergeConfig;
import com.excelconfig.model.StyleConfig;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import com.excelconfig.util.CellValueUtil;
import com.excelconfig.util.StyleCache;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 填充表格策略（FILL_TABLE 模式）
 *
 * <p>支持：
 * <ul>
 *   <li>表头填充</li>
 *   <li>数据行填充</li>
 *   <li>列级智能合并（按数据值纵向合并相同单元格）</li>
 *   <li>跨列合并（colSpan 使单列占据多列宽度）</li>
 *   <li>样式应用（隔行换色、自动列宽）</li>
 * </ul>
 *
 * <p>合并配置优先级：ColumnConfig.merge > ExportConfig.merge（列级覆盖导出级）
 */
class FillTableStrategy implements FillStrategy {

    private static final Logger log = LoggerFactory.getLogger(FillTableStrategy.class);

    @Override
    public void fill(Workbook workbook, ExportConfig config, FillContext context) {
        Sheet sheet = workbook.getSheetAt(0);
        int startRow = context.getStartRow() - 1;
        int startColumn = context.getStartColumn();

        Object data = context.getData().get(config.getKey());
        if (data == null || !(data instanceof List)) {
            return;
        }

        List<?> dataList = (List<?>) data;
        if (dataList.isEmpty()) {
            return;
        }

        List<ColumnConfig> columns = config.getColumns();
        if (columns == null || columns.isEmpty()) {
            return;
        }

        StyleCache styleCache = new StyleCache(workbook);

        // 计算每列物理起始位置（考虑 colspan 偏移）
        int[] colPositions = new int[columns.size()];
        for (int i = 0; i < columns.size(); i++) {
            colPositions[i] = i == 0 ? 0 : colPositions[i - 1] + getColumnSpan(columns.get(i - 1));
        }

        // 1. 填充表头
        fillHeader(sheet, startRow, startColumn, columns, colPositions, config, styleCache);

        // 2. 填充数据行
        fillDataRows(sheet, startRow + 1, startColumn, dataList, columns, colPositions, styleCache);

        // 3. 列级智能合并（方案A）：相同值自动纵向合并
        applyColumnSmartMerge(sheet, startRow + 1, startColumn, dataList, columns, colPositions, config);

        // 4. 应用样式（隔行换色、自动列宽）
        int lastDataRow = startRow + dataList.size();
        applyStyles(sheet, startRow, lastDataRow, startColumn, columns, colPositions, config, styleCache);
    }

    // ==================== 表头填充 ====================

    private void fillHeader(Sheet sheet, int row, int startCol, List<ColumnConfig> columns,
                            int[] colPositions, ExportConfig config, StyleCache styleCache) {
        Row headerRow = getOrCreateRow(sheet, row);

        for (int i = 0; i < columns.size(); i++) {
            ColumnConfig column = columns.get(i);
            int physicalCol = startCol + colPositions[i];
            int span = getColumnSpan(column);

            Cell cell = getOrCreateCell(headerRow, physicalCol);
            cell.setCellValue(column.getHeader() != null ? column.getHeader() : column.getKey());

            // 跨列合并表头（方案C）
            if (span > 1) {
                int mergeEndCol = physicalCol + span - 1;
                CellRangeAddress region = new CellRangeAddress(row, row, physicalCol, mergeEndCol);
                if (!hasOverlappingRegions(sheet, region)) {
                    sheet.addMergedRegion(region);
                    for (int c = physicalCol + 1; c <= mergeEndCol; c++) {
                        Cell covered = getOrCreateCell(headerRow, c);
                        covered.setBlank();
                    }
                }
            }
        }

        // 应用表头样式
        if (config.getHeaderStyle() != null) {
            for (int i = 0; i < columns.size(); i++) {
                int physicalCol = startCol + colPositions[i];
                Cell cell = headerRow.getCell(physicalCol);
                if (cell != null) {
                    applyStyle(cell, config.getHeaderStyle(), styleCache);
                }
            }
        }
    }

    // ==================== 数据行填充 ====================

    private void fillDataRows(Sheet sheet, int startRow, int startCol, List<?> dataList,
                              List<ColumnConfig> columns, int[] colPositions, StyleCache styleCache) {
        for (int i = 0; i < dataList.size(); i++) {
            Object rowObject = dataList.get(i);
            Row row = getOrCreateRow(sheet, startRow + i);
            Map<?, ?> rowMap = rowObject instanceof Map ? (Map<?, ?>) rowObject : null;

            for (int j = 0; j < columns.size(); j++) {
                ColumnConfig column = columns.get(j);
                int physicalCol = startCol + colPositions[j];
                int span = getColumnSpan(column);

                Cell cell = getOrCreateCell(row, physicalCol);
                Object value = getValueFromObject(rowMap, rowObject, column.getKey());
                fillCell(cell, value, column, styleCache);

                // 跨列合并数据单元格（方案C）
                if (span > 1) {
                    int mergeEndCol = physicalCol + span - 1;
                    CellRangeAddress region = new CellRangeAddress(
                        startRow + i, startRow + i, physicalCol, mergeEndCol);
                    if (!hasOverlappingRegions(sheet, region)) {
                        sheet.addMergedRegion(region);
                        for (int c = physicalCol + 1; c <= mergeEndCol; c++) {
                            Cell covered = getOrCreateCell(row, c);
                            covered.setBlank();
                        }
                    }
                }
            }
        }
    }

    // ==================== 列级智能合并（方案A） ====================

    private void applyColumnSmartMerge(Sheet sheet, int startRow, int startCol,
                                       List<?> dataList, List<ColumnConfig> columns,
                                       int[] colPositions, ExportConfig exportConfig) {
        for (int j = 0; j < columns.size(); j++) {
            ColumnConfig column = columns.get(j);
            MergeConfig merge = getEffectiveMerge(column, exportConfig);
            if (merge == null || !merge.isSmartMerge()) {
                continue;
            }

            int physicalCol = startCol + colPositions[j];
            int span = getColumnSpan(column);
            int minSpan = merge.getMinSpan() != null ? merge.getMinSpan() : 2;
            int maxSpan = merge.getMaxSpan() != null ? merge.getMaxSpan() : Integer.MAX_VALUE;

            // 提取该列所有数据值
            List<Object> columnValues = extractColumnValues(dataList, column.getKey());

            // 找连续相同值区间
            List<MergeRange> ranges = findContinuousSameValueRanges(columnValues, minSpan, maxSpan);

            // 创建合并区域
            for (MergeRange range : ranges) {
                int mergeStartRow = startRow + range.startIndex;
                int mergeEndRow = startRow + range.endIndex;
                int mergeEndCol = physicalCol + span - 1;

                CellRangeAddress region = new CellRangeAddress(
                    mergeStartRow, mergeEndRow, physicalCol, mergeEndCol);

                if (!hasOverlappingRegions(sheet, region)) {
                    sheet.addMergedRegion(region);
                    // 清除合并区域内（除左上角外）的所有单元格
                    for (int r = mergeStartRow; r <= mergeEndRow; r++) {
                        Row row = sheet.getRow(r);
                        if (row == null) continue;
                        for (int c = physicalCol; c <= mergeEndCol; c++) {
                            if (r == mergeStartRow && c == physicalCol) continue;
                            Cell cell = row.getCell(c);
                            if (cell != null) {
                                cell.setBlank();
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 从数据行列表中提取指定 key 的所有值
     */
    private List<Object> extractColumnValues(List<?> dataList, String key) {
        List<Object> values = new ArrayList<>(dataList.size());
        for (Object rowObj : dataList) {
            if (rowObj instanceof Map) {
                values.add(((Map<?, ?>) rowObj).get(key));
            } else {
                values.add(null);
            }
        }
        return values;
    }

    /**
     * 获取生效的合并配置：列级 > 导出级
     */
    private MergeConfig getEffectiveMerge(ColumnConfig column, ExportConfig exportConfig) {
        if (column.getMerge() != null) {
            return column.getMerge();
        }
        return exportConfig.getMerge();
    }

    /**
     * 获取列的实际占据宽度（colSpan 默认为 1）
     */
    private int getColumnSpan(ColumnConfig column) {
        if (column.getMerge() != null && column.getMerge().getColSpan() != null) {
            int span = column.getMerge().getColSpan();
            return span > 0 ? span : 1;
        }
        return 1;
    }

    // ==================== 区间查找 ====================

    /**
     * 找出连续相同值的区间
     */
    private List<MergeRange> findContinuousSameValueRanges(List<?> values, int minSpan, int maxSpan) {
        List<MergeRange> ranges = new ArrayList<>();
        int n = values.size();
        int start = 0;
        while (start < n) {
            Object current = values.get(start);
            int end = start;
            while (end + 1 < n && isSameValue(current, values.get(end + 1))) {
                end++;
            }
            int span = end - start + 1;
            if (span >= minSpan && span <= maxSpan) {
                ranges.add(new MergeRange(start, end));
            }
            start = end + 1;
        }
        return ranges;
    }

    private boolean isSameValue(Object v1, Object v2) {
        if (v1 == null && v2 == null) return true;
        if (v1 == null || v2 == null) return false;
        return v1.equals(v2);
    }

    // ==================== 样式应用 ====================

    private void applyStyles(Sheet sheet, int headerRow, int lastDataRow, int startCol,
                             List<ColumnConfig> columns, int[] colPositions,
                             ExportConfig config, StyleCache styleCache) {
        // 隔行换色
        if (config.getAlternateRows() != null && config.getAlternateRows()) {
            StyleConfig alternateStyle = new StyleConfig();
            alternateStyle.setBackground("#D9D9D9");
            CellStyle cachedStyle = styleCache.getOrCreateStyle(alternateStyle);

            for (int i = headerRow + 1; i <= lastDataRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                if (i % 2 == 0) {
                    for (int j = 0; j < columns.size(); j++) {
                        int physicalCol = startCol + colPositions[j];
                        int span = getColumnSpan(columns.get(j));
                        for (int c = physicalCol; c < physicalCol + span; c++) {
                            Cell cell = row.getCell(c);
                            if (cell == null) continue;
                            cell.setCellStyle(cachedStyle);
                        }
                    }
                }
            }
        }

        // 自动列宽
        if (config.getAutoWidth() != null && config.getAutoWidth()) {
            for (int i = 0; i < columns.size(); i++) {
                ColumnConfig column = columns.get(i);
                if (column.getWidth() != null) {
                    int physicalCol = startCol + colPositions[i];
                    sheet.setColumnWidth(physicalCol, column.getWidth() * 256);
                }
            }
        }
    }

    // ==================== 单元格操作 ====================

    private void fillCell(Cell cell, Object value, ColumnConfig column, StyleCache styleCache) {
        CellValueUtil.setCellValue(cell, value);

        String format = (column.getFormat() != null && value instanceof Number) ? column.getFormat() : null;
        StyleConfig effectiveStyle = mergeStyleAndFormat(column.getStyle(), format);
        if (effectiveStyle != null) {
            CellStyle cachedStyle = styleCache.getOrCreateStyle(effectiveStyle);
            if (cachedStyle != null) {
                cell.setCellStyle(cachedStyle);
            }
        }
    }

    private static StyleConfig mergeStyleAndFormat(StyleConfig style, String format) {
        if (format == null) return style;
        if (style == null) {
            StyleConfig fmtOnly = new StyleConfig();
            fmtOnly.setFormat(format);
            return fmtOnly;
        }
        StyleConfig merged = new StyleConfig();
        merged.setBold(style.getBold());
        merged.setBackground(style.getBackground());
        merged.setFontColor(style.getFontColor());
        merged.setFontSize(style.getFontSize());
        merged.setHorizontalAlign(style.getHorizontalAlign());
        merged.setVerticalAlign(style.getVerticalAlign());
        merged.setFormat(format);
        return merged;
    }

    private void applyStyle(Cell cell, StyleConfig style, StyleCache styleCache) {
        CellStyle cellStyle = styleCache.getOrCreateStyle(style);
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    // ==================== 行/列操作 ====================

    private Row getOrCreateRow(Sheet sheet, int rowNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
    }

    private Cell getOrCreateCell(Row row, int column) {
        Cell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        return cell;
    }

    private Object getValueFromObject(Map<?, ?> map, Object obj, String key) {
        if (map != null) {
            return map.get(key);
        }
        return null;
    }

    // ==================== 合并区域重叠检测 ====================

    private boolean hasOverlappingRegions(Sheet sheet, CellRangeAddress newRegion) {
        for (CellRangeAddress existing : sheet.getMergedRegions()) {
            if (regionsOverlap(existing, newRegion)) {
                return true;
            }
        }
        return false;
    }

    private boolean regionsOverlap(CellRangeAddress r1, CellRangeAddress r2) {
        return !(r1.getLastRow() < r2.getFirstRow()
                || r1.getFirstRow() > r2.getLastRow()
                || r1.getLastColumn() < r2.getFirstColumn()
                || r1.getFirstColumn() > r2.getLastColumn());
    }

    // ==================== 内部类 ====================

    /**
     * 合并区间（起始偏移量 + 结束偏移量，相对于数据列表）
     */
    private static class MergeRange {
        final int startIndex;
        final int endIndex;

        MergeRange(int startIndex, int endIndex) {
            this.startIndex = startIndex;
            this.endIndex = endIndex;
        }
    }

    @Override
    public com.excelconfig.spi.FillMode getSupportedMode() {
        return com.excelconfig.spi.FillMode.FILL_TABLE;
    }
}
