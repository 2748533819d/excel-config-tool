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
 * 提取到空行停止策略（UNTIL_EMPTY 模式）
 */
public class UntilEmptyExtractStrategy implements ExtractStrategy {

    private final CellParser cellParser;

    public UntilEmptyExtractStrategy() {
        this.cellParser = new DefaultCellParser();
    }

    @Override
    public List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context) {
        List<Object> result = new ArrayList<>();

        int startRow = context.getStartRow();
        int column = context.getStartColumn();

        ParserConfig parserConfig = config.getParser();

        for (int rowNum = startRow; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);

            if (row == null) {
                break;  // 遇到空行停止
            }

            Cell cell = row.getCell(column);
            if (cell == null || isCellEmpty(cell)) {
                break;  // 遇到空单元格停止
            }

            Object value = cellParser.parse(cell, parserConfig);
            result.add(value != null ? value : "");
        }

        return result;
    }

    @Override
    public com.excelconfig.spi.ExtractMode getSupportedMode() {
        return com.excelconfig.spi.ExtractMode.UNTIL_EMPTY;
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
