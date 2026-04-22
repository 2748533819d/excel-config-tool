package com.excelconfig.extract;

import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.ParserConfig;
import com.excelconfig.spi.CellParser;
import com.excelconfig.spi.ExtractContext;
import com.excelconfig.spi.ExtractStrategy;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * 单个单元格提取策略（SINGLE 模式）
 */
public class SingleExtractStrategy implements ExtractStrategy {

    private final CellParser cellParser;

    public SingleExtractStrategy() {
        this.cellParser = new DefaultCellParser();
    }

    @Override
    public List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context) {
        Row row = sheet.getRow(context.getStartRow());
        if (row == null) {
            return Collections.emptyList();
        }

        Cell cell = row.getCell(context.getStartColumn());
        if (cell == null) {
            return Collections.emptyList();
        }

        ParserConfig parserConfig = config.getParser();
        Object value = cellParser.parse(cell, parserConfig);

        return value != null ? Collections.singletonList(value) : Collections.emptyList();
    }

    @Override
    public com.excelconfig.spi.ExtractMode getSupportedMode() {
        return com.excelconfig.spi.ExtractMode.SINGLE;
    }
}
