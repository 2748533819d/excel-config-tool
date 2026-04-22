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
 * 区域提取策略（BLOCK 模式）
 */
public class BlockExtractStrategy implements ExtractStrategy {

    private final CellParser cellParser;

    public BlockExtractStrategy() {
        this.cellParser = new DefaultCellParser();
    }

    @Override
    public List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context) {
        List<Object> result = new ArrayList<>();

        int startRow = context.getStartRow();
        int startCol = context.getStartColumn();
        Integer endRow = context.getEndRow();
        Integer endCol = context.getEndColumn();

        // 如果没有指定结束位置，使用固定行数/列数
        if (endRow == null && config.getRange() != null && config.getRange().getRows() != null) {
            endRow = startRow + config.getRange().getRows() - 1;
        }
        if (endCol == null && config.getRange() != null && config.getRange().getCols() != null) {
            endCol = startCol + config.getRange().getCols() - 1;
        }

        // 默认提取 1 行 1 列
        if (endRow == null) endRow = startRow;
        if (endCol == null) endCol = startCol;

        ParserConfig parserConfig = config.getParser();

        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            List<Object> rowData = new ArrayList<>();

            if (row != null) {
                for (int colNum = startCol; colNum <= endCol; colNum++) {
                    Cell cell = row.getCell(colNum);
                    Object value = cell != null ? cellParser.parse(cell, parserConfig) : null;
                    rowData.add(value != null ? value : "");
                }
            } else {
                // 空行，填充空值
                for (int colNum = startCol; colNum <= endCol; colNum++) {
                    rowData.add("");
                }
            }

            result.add(rowData);
        }

        return result;
    }

    @Override
    public com.excelconfig.spi.ExtractMode getSupportedMode() {
        return com.excelconfig.spi.ExtractMode.BLOCK;
    }
}
