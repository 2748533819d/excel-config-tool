package com.excelconfig.export;

import com.excelconfig.model.ExportConfig;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import com.excelconfig.util.CellValueUtil;
import com.excelconfig.util.StyleCache;
import org.apache.poi.ss.usermodel.*;

/**
 * 填充单个单元格策略（FILL_CELL 模式）
 */
class FillCellStrategy implements FillStrategy {

    @Override
    public void fill(Workbook workbook, ExportConfig config, FillContext context) {
        Sheet sheet = workbook.getSheetAt(0);
        int row = context.getStartRow();
        int column = context.getStartColumn();

        Row targetRow = getOrCreateRow(sheet, row);
        Cell targetCell = getOrCreateCell(targetRow, column);

        // 获取数据
        Object data = context.getData().get(config.getKey());
        fillCell(targetCell, data, config.getStyle(), workbook);
    }

    /**
     * 填充单元格
     */
    protected void fillCell(Cell cell, Object value, com.excelconfig.model.StyleConfig style) {
        fillCell(cell, value, style, cell.getSheet().getWorkbook());
    }

    protected void fillCell(Cell cell, Object value, com.excelconfig.model.StyleConfig style, Workbook workbook) {
        fillCell(cell, value, style, new StyleCache(workbook));
    }

    /**
     * 填充单元格（使用外部传入的 StyleCache，实现 fill() 级别复用）
     */
    protected void fillCell(Cell cell, Object value, com.excelconfig.model.StyleConfig style, StyleCache styleCache) {
        CellValueUtil.setCellValue(cell, value);

        // 应用样式（使用缓存）
        if (style != null) {
            CellStyle cellStyle = styleCache.getOrCreateStyle(style);
            if (cellStyle != null) {
                cell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 应用样式（使用缓存）
     */
    protected void applyStyle(Cell cell, com.excelconfig.model.StyleConfig style) {
        applyStyle(cell, style, cell.getSheet().getWorkbook());
    }

    protected void applyStyle(Cell cell, com.excelconfig.model.StyleConfig style, Workbook workbook) {
        applyStyle(cell, style, new StyleCache(workbook));
    }

    protected void applyStyle(Cell cell, com.excelconfig.model.StyleConfig style, StyleCache styleCache) {
        CellStyle cellStyle = styleCache.getOrCreateStyle(style);
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    protected Row getOrCreateRow(Sheet sheet, int rowNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
    }

    protected Cell getOrCreateCell(Row row, int column) {
        Cell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        return cell;
    }

    @Override
    public com.excelconfig.spi.FillMode getSupportedMode() {
        return com.excelconfig.spi.FillMode.FILL_CELL;
    }
}
