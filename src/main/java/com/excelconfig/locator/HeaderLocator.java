package com.excelconfig.locator;

import com.excelconfig.model.HeaderConfig;
import org.apache.poi.ss.usermodel.*;

/**
 * 表头定位器
 *
 * 核心功能：
 * 1. 根据表头文字匹配定位单元格
 * 2. 支持全局搜索（默认）
 * 3. 支持指定行范围搜索（可选）
 */
public class HeaderLocator {

    /**
     * 定位表头
     *
     * @param sheet Excel Sheet
     * @param config 表头配置
     * @return 表头位置
     * @throws HeaderNotFoundException 当未找到匹配的表头时
     */
    public HeaderPosition locate(Sheet sheet, HeaderConfig config) {
        if (config.getMatch() == null || config.getMatch().isEmpty()) {
            throw new IllegalArgumentException("表头匹配文本不能为空");
        }

        // 确定搜索范围
        int startRow = 0;
        int endRow = sheet.getLastRowNum();

        if (config.getInRows() != null && config.getInRows().length == 2) {
            // 指定范围搜索（转换为 0-based 索引）
            startRow = config.getInRows()[0] - 1;
            endRow = config.getInRows()[1] - 1;
        }

        // 在范围内搜索
        for (int rowNum = startRow; rowNum <= endRow && rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }

            for (Cell cell : row) {
                if (cell == null) {
                    continue;
                }

                String cellValue = getCellValueAsString(cell);
                if (cellValue != null && cellValue.equals(config.getMatch())) {
                    return new HeaderPosition(rowNum, cell.getColumnIndex());
                }
            }
        }

        throw new HeaderNotFoundException(
            String.format("未找到表头 '%s' (搜索范围：行 %d-%d)",
                config.getMatch(), startRow + 1, endRow + 1));
    }

    /**
     * 获取单元格的字符串值
     */
    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double num = cell.getNumericCellValue();
                    if (num == (long) num) {
                        return String.valueOf((long) num);
                    }
                    return String.valueOf(num);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            default:
                return null;
        }
    }
}
