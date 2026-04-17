package com.excelconfig.spi;

import org.apache.poi.ss.usermodel.Cell;
import com.excelconfig.model.ParserConfig;

/**
 * 单元格解析器接口
 */
public interface CellParser {

    /**
     * 解析单元格
     * @param cell 单元格
     * @param config 解析器配置
     * @return 解析后的值
     */
    Object parse(Cell cell, ParserConfig config);
}
