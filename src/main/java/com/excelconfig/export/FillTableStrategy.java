package com.excelconfig.export;

import com.excelconfig.model.ColumnConfig;
import com.excelconfig.model.ExportConfig;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * 填充表格策略（FILL_TABLE 模式）
 *
 * 支持：
 * 1. 表头填充
 * 2. 数据行填充
 * 3. 样式应用（隔行换色、条件格式等）
 */
class FillTableStrategy implements FillStrategy {

    @Override
    public void fill(Workbook workbook, ExportConfig config, FillContext context) {
        Sheet sheet = workbook.getSheetAt(0);
        // 对于 FILL_TABLE 模式，context.getStartRow() 是表头下方一行，所以表头从 startRow - 1 开始
        int startRow = context.getStartRow() - 1;
        int startColumn = context.getStartColumn();

        // 获取数据
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

        // 1. 填充表头
        fillHeader(sheet, startRow, startColumn, columns, config);

        // 2. 填充数据行
        fillDataRows(sheet, startRow + 1, startColumn, dataList, columns, config);

        // 3. 应用样式
        applyStyles(sheet, startRow, startRow + dataList.size(), startColumn, columns, config);
    }

    /**
     * 填充表头
     */
    private void fillHeader(Sheet sheet, int row, int startCol, List<ColumnConfig> columns,
                           ExportConfig config) {
        Row headerRow = getOrCreateRow(sheet, row);

        for (int i = 0; i < columns.size(); i++) {
            ColumnConfig column = columns.get(i);
            Cell cell = getOrCreateCell(headerRow, startCol + i);
            cell.setCellValue(column.getHeader() != null ? column.getHeader() : column.getKey());
        }

        // 应用表头样式
        if (config.getHeaderStyle() != null) {
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = headerRow.getCell(startCol + i);
                applyStyle(cell, config.getHeaderStyle());
            }
        }
    }

    /**
     * 填充数据行
     */
    private void fillDataRows(Sheet sheet, int startRow, int startCol, List<?> dataList,
                              List<ColumnConfig> columns, ExportConfig config) {
        for (int i = 0; i < dataList.size(); i++) {
            Object rowObject = dataList.get(i);
            Row row = getOrCreateRow(sheet, startRow + i);

            Map<?, ?> rowMap = rowObject instanceof Map ? (Map<?, ?>) rowObject : null;

            for (int j = 0; j < columns.size(); j++) {
                ColumnConfig column = columns.get(j);
                Cell cell = getOrCreateCell(row, startCol + j);

                Object value = getValueFromObject(rowMap, rowObject, column.getKey());
                fillCell(cell, value, column);
            }
        }
    }

    /**
     * 从对象中获取字段值
     */
    private Object getValueFromObject(Map<?, ?> map, Object obj, String key) {
        if (map != null) {
            return map.get(key);
        }
        // 支持通过 getter 获取值
        return null;
    }

    /**
     * 填充单元格
     */
    private void fillCell(Cell cell, Object value, ColumnConfig column) {
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

        // 应用列样式
        if (column.getStyle() != null) {
            applyStyle(cell, column.getStyle());
        }

        // 应用数字格式（使用新创建的 CellStyle 避免修改共享样式）
        if (column.getFormat() != null && value instanceof Number) {
            Workbook wb = cell.getSheet().getWorkbook();
            DataFormat format = wb.createDataFormat();
            CellStyle style = wb.createCellStyle();
            style.setDataFormat(format.getFormat(column.getFormat()));
            cell.setCellStyle(style);
        }
    }

    /**
     * 应用样式
     */
    private void applyStyles(Sheet sheet, int headerRow, int lastDataRow, int startCol,
                            List<ColumnConfig> columns, ExportConfig config) {
        // 隔行换色
        if (config.getAlternateRows() != null && config.getAlternateRows()) {
            for (int i = headerRow + 1; i <= lastDataRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                for (int j = 0; j < columns.size(); j++) {
                    Cell cell = row.getCell(startCol + j);
                    if (cell == null) continue;

                    if (i % 2 == 0) {
                        CellStyle style = sheet.getWorkbook().createCellStyle();
                        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        cell.setCellStyle(style);
                    }
                }
            }
        }

        // 自动列宽
        if (config.getAutoWidth() != null && config.getAutoWidth()) {
            for (int i = 0; i < columns.size(); i++) {
                ColumnConfig column = columns.get(i);
                if (column.getWidth() != null) {
                    sheet.setColumnWidth(startCol + i, column.getWidth() * 256);
                }
            }
        }
    }

    private void applyStyle(Cell cell, com.excelconfig.model.StyleConfig style) {
        Workbook wb = cell.getSheet().getWorkbook();
        // POI 的 CellStyle 在 Workbook 级别共享，必须创建新实例而非修改 getCellStyle()
        CellStyle cellStyle = wb.createCellStyle();

        // 加粗
        if (style.getBold() != null && style.getBold()) {
            Font font = wb.createFont();
            font.setBold(true);
            cellStyle.setFont(font);
        }

        if (style.getBackground() != null) {
            applyBackgroundColor(cell, cellStyle, style.getBackground());
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

        cell.setCellStyle(cellStyle);
    }

    private void applyBackgroundColor(Cell cell, CellStyle cellStyle, String colorHex) {
        Workbook wb = cell.getSheet().getWorkbook();
        if (wb instanceof XSSFWorkbook) {
            try {
                String hex = colorHex.startsWith("#") ? colorHex.substring(1) : colorHex;
                if (hex.length() == 6) {
                    int r = Integer.parseInt(hex.substring(0, 2), 16);
                    int g = Integer.parseInt(hex.substring(2, 4), 16);
                    int b = Integer.parseInt(hex.substring(4, 6), 16);
                    XSSFColor xssfColor = new XSSFColor(new java.awt.Color(r, g, b), new DefaultIndexedColorMap());
                    cellStyle.setFillForegroundColor(xssfColor);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    return;
                }
            } catch (Exception ignored) {
            }
        }
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

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

    @Override
    public com.excelconfig.spi.FillMode getSupportedMode() {
        return com.excelconfig.spi.FillMode.FILL_TABLE;
    }
}
