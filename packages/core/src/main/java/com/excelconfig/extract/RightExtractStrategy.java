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
 * 向右提取策略（RIGHT 模式）
 */
public class RightExtractStrategy implements ExtractStrategy {

    private final CellParser cellParser;

    public RightExtractStrategy() {
        this.cellParser = new DefaultCellParser();
    }

    @Override
    public List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context) {
        List<Object> result = new ArrayList<>();

        int rowNum = context.getStartRow();
        int startColumn = context.getStartColumn();

        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return result;
        }

        Boolean skipEmpty = config.getRange() != null ? config.getRange().getSkipEmpty() : true;
        Integer maxCols = config.getRange() != null ? config.getRange().getMaxRows() : null;
        Integer fixedCols = config.getRange() != null ? config.getRange().getCols() : null;

        ParserConfig parserConfig = config.getParser();
        int consecutiveEmptyCells = 0;

        for (int colNum = startColumn; ; colNum++) {
            if (maxCols != null && result.size() >= maxCols) {
                break;
            }

            if (fixedCols != null && result.size() >= fixedCols) {
                break;
            }

            Cell cell = row.getCell(colNum);

            if (cell == null || isCellEmpty(cell)) {
                if (skipEmpty) {
                    consecutiveEmptyCells++;
                    if (consecutiveEmptyCells >= 2) {
                        break;
                    }
                    continue;
                }
            } else {
                consecutiveEmptyCells = 0;
                Object value = cellParser.parse(cell, parserConfig);
                if (value != null) {
                    result.add(value);
                }
            }
        }

        return result;
    }

    @Override
    public com.excelconfig.spi.ExtractMode getSupportedMode() {
        return com.excelconfig.spi.ExtractMode.RIGHT;
    }

    private boolean isCellEmpty(Cell cell) {
        switch (cell.getCellType()) {
            case BLANK:
                return true;
            case STRING:
                return cell.getStringCellValue() == null || cell.getStringCellValue().trim().isEmpty();
            default:
                return false;
        }
    }
}
