package com.excelconfig.sax;

import java.util.List;

/**
 * 行数据回调接口
 */
public interface RowHandler {

    /**
     * 处理一行数据
     *
     * @param rowNum 行号（从 0 开始）
     * @param cells 单元格值列表
     */
    void handleRow(int rowNum, List<String> cells);
}
