package com.excelconfig.extract;

import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.ParserConfig;
import com.excelconfig.spi.CellParser;
import com.excelconfig.spi.ExtractContext;
import com.excelconfig.spi.ExtractStrategy;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

/**
 * 向下提取策略（DOWN 模式）
 *
 * 从表头下方开始向下提取数据，直到：
 * 1. 遇到空行（skipEmpty=true）
 * 2. 达到 maxRows 限制
 * 3. 到达 sheet 末尾
 */
public class DownExtractStrategy implements ExtractStrategy {

    private final CellParser cellParser;

    public DownExtractStrategy() {
        this.cellParser = new DefaultCellParser();
    }

    @Override
    public List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context) {
        List<Object> result = new ArrayList<>();

        int startRow = context.getStartRow();
        int column = context.getStartColumn();

        // 获取范围配置
        Boolean skipEmpty = config.getRange() != null ? config.getRange().getSkipEmpty() : true;
        Integer maxRows = config.getRange() != null ? config.getRange().getMaxRows() : null;
        Integer fixedRows = config.getRange() != null ? config.getRange().getRows() : null;

        ParserConfig parserConfig = config.getParser();

        int rowsRead = 0;
        int consecutiveEmptyRows = 0;

        for (int rowNum = startRow; ; rowNum++) {
            // 检查是否达到最大行数限制
            if (maxRows != null && rowsRead >= maxRows) {
                break;
            }

            // 检查是否达到固定行数
            if (fixedRows != null && rowsRead >= fixedRows) {
                break;
            }

            // 检查是否到达 sheet 末尾
            if (rowNum > sheet.getLastRowNum()) {
                break;
            }

            Row row = sheet.getRow(rowNum);
            Object value = null;

            if (row == null) {
                // 空行
                if (skipEmpty) {
                    consecutiveEmptyRows++;
                    if (consecutiveEmptyRows >= 2) {
                        // 连续 2 个空行，停止
                        break;
                    }
                    continue;
                } else {
                    value = null;
                }
            } else {
                Cell cell = row.getCell(column);
                if (cell == null || isCellEmpty(cell)) {
                    if (skipEmpty) {
                        consecutiveEmptyRows++;
                        if (consecutiveEmptyRows >= 2) {
                            break;
                        }
                        continue;
                    } else {
                        value = null;
                    }
                } else {
                    consecutiveEmptyRows = 0;
                    value = cellParser.parse(cell, parserConfig);
                }
            }

            if (value != null) {
                result.add(value);
                rowsRead++;
            }
        }

        return result;
    }

    @Override
    public com.excelconfig.spi.ExtractMode getSupportedMode() {
        return com.excelconfig.spi.ExtractMode.DOWN;
    }

    /**
     * 判断单元格是否为空
     */
    private boolean isCellEmpty(Cell cell) {
        switch (cell.getCellType()) {
            case BLANK:
                return true;
            case STRING:
                return cell.getStringCellValue() == null || cell.getStringCellValue().trim().isEmpty();
            case NUMERIC:
                return false;
            case BOOLEAN:
                return false;
            default:
                return true;
        }
    }
}
