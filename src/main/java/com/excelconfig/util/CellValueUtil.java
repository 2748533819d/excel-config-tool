package com.excelconfig.util;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 单元格值设置工具 - 统一处理 POI 单元格的类型转换
 *
 * <p>集中管理 {@link Cell#setCellValue} 的类型分支逻辑，
 * 新增值类型支持时只需修改此处，避免在多处 Strategy 中重复更新。</p>
 */
public final class CellValueUtil {

    private CellValueUtil() {
    }

    /**
     * 根据值的实际类型设置单元格内容
     *
     * @param cell  目标单元格
     * @param value 待写入的值（支持 String、Number、Boolean、Date，其他类型调用 toString）
     */
    public static void setCellValue(Cell cell, Object value) {
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
    }
}
