package com.excelconfig.export;

import com.excelconfig.model.ExportConfig;
import com.excelconfig.model.MergeConfig;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * 向下填充策略（FILL_DOWN 模式）
 *
 * 核心功能：
 * 1. 根据数据量自动填充
 * 2. 检查下方是否有其他表，如有需要则下移
 * 3. 支持合并单元格（智能合并和固定区域合并）
 */
public class FillDownStrategy extends FillCellStrategy implements FillStrategy {

    @Override
    public void fill(Workbook workbook, ExportConfig config, FillContext context) {
        Sheet sheet = workbook.getSheetAt(0);
        int startRow = context.getStartRow();
        int column = context.getStartColumn();

        // 获取数据
        Object data = context.getData().get(config.getKey());
        if (data == null) {
            return;
        }

        List<?> dataList = convertToList(data);
        if (dataList.isEmpty()) {
            return;
        }

        // 检查是否需要下移下方内容
        Integer maxRows = config.getMaxRows();
        int fillRows = maxRows != null ? Math.min(dataList.size(), maxRows) : dataList.size();

        // 先下移填充区域内的原有数据
        shiftExistingDataInRange(sheet, startRow, fillRows, column);

        // 检查是否有合并配置
        MergeConfig merge = config.getMerge();
        if (merge != null && merge.isSmartMerge()) {
            // 智能合并模式：按数据值合并相同值的单元格
            fillWithSmartMerge(sheet, dataList, fillRows, startRow, column, config.getStyle(), merge);
        } else if (merge != null && merge.isFixedMerge()) {
            // 固定区域合并模式
            fillWithFixedMerge(sheet, dataList, fillRows, startRow, column, config.getStyle(), merge);
        } else {
            // 普通填充模式
            fillNormal(sheet, dataList, fillRows, startRow, column, config.getStyle());
        }
    }

    /**
     * 普通填充模式（无合并）
     */
    private void fillNormal(Sheet sheet, List<?> dataList, int fillRows, int startRow, int column,
                            com.excelconfig.model.StyleConfig style) {
        for (int i = 0; i < fillRows; i++) {
            Row row = getOrCreateRow(sheet, startRow + i);
            Cell cell = getOrCreateCell(row, column);
            fillCell(cell, dataList.get(i), style);
        }
    }

    /**
     * 智能合并模式：按数据值自动合并相同值的单元格
     *
     * 算法：
     * 1. 遍历数据，找出连续相同值的区间
     * 2. 对每个区间创建合并区域
     * 3. 只在区间的第一个单元格填入值
     */
    private void fillWithSmartMerge(Sheet sheet, List<?> dataList, int fillRows, int startRow, int column,
                                    com.excelconfig.model.StyleConfig style, MergeConfig merge) {
        int minSpan = merge.getMinSpan() != null ? merge.getMinSpan() : 2;
        int maxSpan = merge.getMaxSpan() != null ? merge.getMaxSpan() : Integer.MAX_VALUE;

        // 找出连续相同值的区间
        List<MergeRange> mergeRanges = findContinuousSameValueRanges(dataList, fillRows, minSpan, maxSpan);

        // 先填充所有数据（不合并）
        for (int i = 0; i < fillRows; i++) {
            Row row = getOrCreateRow(sheet, startRow + i);
            Cell cell = getOrCreateCell(row, column);
            fillCell(cell, dataList.get(i), style);
        }

        // 再创建合并区域
        for (MergeRange range : mergeRanges) {
            int mergeStartRow = startRow + range.startIndex;
            int mergeEndRow = startRow + range.endIndex;

            // 创建合并区域
            CellRangeAddress region = new CellRangeAddress(mergeStartRow, mergeEndRow, column, column);

            // 检查是否已与现有合并区域重叠
            if (!hasOverlappingRegions(sheet, region)) {
                sheet.addMergedRegion(region);

                // 清除合并区域内除第一个外的所有单元格
                for (int r = mergeStartRow + 1; r <= mergeEndRow; r++) {
                    Row row = getOrCreateRow(sheet, r);
                    Cell cell = getOrCreateCell(row, column);
                    cell.setBlank();
                }
            }
        }
    }

    /**
     * 找出连续相同值的区间
     *
     * @param dataList 数据列表
     * @param fillRows 填充行数
     * @param minSpan 最小合并数量
     * @param maxSpan 最大合并数量
     * @return 需要合并的区间列表
     */
    private List<MergeRange> findContinuousSameValueRanges(List<?> dataList, int fillRows, int minSpan, int maxSpan) {
        List<MergeRange> ranges = new ArrayList<>();
        int startIndex = 0;

        while (startIndex < fillRows) {
            Object currentValue = dataList.get(startIndex);
            int endIndex = startIndex;

            // 向后查找相同值的范围
            while (endIndex + 1 < fillRows) {
                Object nextValue = dataList.get(endIndex + 1);
                if (isSameValue(currentValue, nextValue)) {
                    endIndex++;
                } else {
                    break;
                }
            }

            // 检查是否满足合并条件
            int span = endIndex - startIndex + 1;
            if (span >= minSpan && span <= maxSpan) {
                ranges.add(new MergeRange(startIndex, endIndex));
            }

            // 移动到下一个区间
            startIndex = endIndex + 1;
        }

        return ranges;
    }

    /**
     * 判断两个值是否相同
     */
    private boolean isSameValue(Object v1, Object v2) {
        if (v1 == null && v2 == null) {
            return true;
        }
        if (v1 == null || v2 == null) {
            return false;
        }
        return v1.equals(v2);
    }

    /**
     * 固定区域合并模式
     *
     * 每个数据按 rowSpan/colSpan 指定的大小合并，起始位置累加
     */
    private void fillWithFixedMerge(Sheet sheet, List<?> dataList, int fillRows, int startRow, int column,
                                    com.excelconfig.model.StyleConfig style, MergeConfig merge) {
        int rowSpan = merge.getRowSpan() != null ? merge.getRowSpan() : 1;
        int colSpan = merge.getColSpan() != null ? merge.getColSpan() : 1;
        int startRowOffset = merge.getStartRowOffset() != null ? merge.getStartRowOffset() : 0;
        int startColOffset = merge.getStartColOffset() != null ? merge.getStartColOffset() : 0;

        for (int i = 0; i < fillRows; i++) {
            // 计算当前数据的起始位置（累加 rowSpan）
            int mergeStartRow = startRow + startRowOffset + (i * rowSpan);
            int mergeStartCol = column + startColOffset;
            int mergeEndRow = mergeStartRow + rowSpan - 1;
            int mergeEndCol = mergeStartCol + colSpan - 1;

            // 创建合并区域
            CellRangeAddress region = new CellRangeAddress(mergeStartRow, mergeEndRow, mergeStartCol, mergeEndCol);

            if (!hasOverlappingRegions(sheet, region)) {
                sheet.addMergedRegion(region);

                // 填充数据到合并区域的第一个单元格
                Row row = getOrCreateRow(sheet, mergeStartRow);
                Cell cell = getOrCreateCell(row, mergeStartCol);
                fillCell(cell, dataList.get(i), style);

                // 清除合并区域内的其他单元格
                clearMergedCells(sheet, mergeStartRow, mergeEndRow, mergeStartCol, mergeEndCol);
            }
        }
    }

    /**
     * 检查是否与现有合并区域重叠
     */
    private boolean hasOverlappingRegions(Sheet sheet, CellRangeAddress newRegion) {
        for (CellRangeAddress existing : sheet.getMergedRegions()) {
            if (regionsOverlap(existing, newRegion)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 检查两个合并区域是否重叠
     */
    private boolean regionsOverlap(CellRangeAddress r1, CellRangeAddress r2) {
        return !(r1.getLastRow() < r2.getFirstRow() ||
                r1.getFirstRow() > r2.getLastRow() ||
                r1.getLastColumn() < r2.getFirstColumn() ||
                r1.getFirstColumn() > r2.getLastColumn());
    }

    /**
     * 清除合并区域内的其他单元格
     */
    private void clearMergedCells(Sheet sheet, int startRow, int endRow, int startCol, int endCol) {
        for (int r = startRow; r <= endRow; r++) {
            for (int c = startCol; c <= endCol; c++) {
                // 跳过左上角单元格（它保存实际数据）
                if (r == startRow && c == startCol) {
                    continue;
                }
                Row row = sheet.getRow(r);
                if (row != null) {
                    Cell cell = row.getCell(c);
                    if (cell != null) {
                        cell.setBlank();
                    }
                }
            }
        }
    }

    /**
     * 下移填充区域内的原有数据
     */
    private void shiftExistingDataInRange(Sheet sheet, int startRow, int fillRows, int column) {
        int endRow = startRow + fillRows - 1;

        // 从下往上遍历，找出范围内的有内容的行并下移
        for (int rowNum = endRow; rowNum >= startRow; rowNum--) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cell = row.getCell(column);
                if (cell != null && !isCellEmpty(cell, column)) {
                    Row newRow = getOrCreateRow(sheet, rowNum + fillRows);
                    Cell newCell = getOrCreateCell(newRow, column);
                    copyCellValue(cell, newCell);
                }
            }
        }
    }

    /**
     * 复制单元格的值（不复制样式）
     */
    private void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }

    /**
     * 判断单元格是否为空
     */
    private boolean isCellEmpty(Cell cell, int column) {
        switch (cell.getCellType()) {
            case BLANK:
                return true;
            case STRING:
                return cell.getStringCellValue() == null || cell.getStringCellValue().trim().isEmpty();
            default:
                return false;
        }
    }

    /**
     * 将对象转换为 List
     */
    @SuppressWarnings("unchecked")
    private List<?> convertToList(Object data) {
        if (data instanceof List) {
            return (List<?>) data;
        } else if (data instanceof Collection) {
            return (List<?>) data;
        } else if (data.getClass().isArray()) {
            Object[] array = (Object[]) data;
            return java.util.Arrays.asList(array);
        } else {
            return java.util.Collections.singletonList(data);
        }
    }

    /**
     * 合并区间
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
        return com.excelconfig.spi.FillMode.FILL_DOWN;
    }
}
