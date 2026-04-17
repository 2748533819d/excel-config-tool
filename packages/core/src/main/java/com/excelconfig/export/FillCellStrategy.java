package com.excelconfig.export;

import com.excelconfig.model.ExportConfig;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import org.apache.poi.ss.usermodel.*;

import java.util.Collections;
import java.util.List;

/**
 * 填充单个单元格策略（FILL_CELL 模式）
 */
public class FillCellStrategy implements FillStrategy {

    @Override
    public void fill(Workbook workbook, ExportConfig config, FillContext context) {
        Sheet sheet = workbook.getSheetAt(0);
        int row = context.getStartRow();
        int column = context.getStartColumn();

        Row targetRow = getOrCreateRow(sheet, row);
        Cell targetCell = getOrCreateCell(targetRow, column);

        // 获取数据
        Object data = context.getData().get(config.getKey());
        fillCell(targetCell, data, config.getStyle());
    }

    /**
     * 填充单元格
     */
    protected void fillCell(Cell cell, Object value, com.excelconfig.model.StyleConfig style) {
        if (value == null) {
            cell.setBlank();
            return;
        }

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof java.util.Date) {
            cell.setCellValue((java.util.Date) value);
        } else {
            cell.setCellValue(value.toString());
        }

        // 应用样式
        if (style != null) {
            applyStyle(cell, style);
        }
    }

    /**
     * 应用样式
     */
    protected void applyStyle(Cell cell, com.excelconfig.model.StyleConfig style) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        }

        // 水平对齐
        if (style.getHorizontalAlign() != null) {
            switch (style.getHorizontalAlign().toUpperCase()) {
                case "LEFT":
                    cellStyle.setAlignment(HorizontalAlignment.LEFT);
                    break;
                case "CENTER":
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case "RIGHT":
                    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                    break;
            }
        }

        // 垂直对齐
        if (style.getVerticalAlign() != null) {
            switch (style.getVerticalAlign().toUpperCase()) {
                case "TOP":
                    cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
                    break;
                case "CENTER":
                    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    break;
                case "BOTTOM":
                    cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    break;
            }
        }

        // 数字格式
        if (style.getFormat() != null) {
            cellStyle.setDataFormat(cell.getSheet().getWorkbook().createDataFormat().getFormat(style.getFormat()));
        }

        cell.setCellStyle(cellStyle);
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
