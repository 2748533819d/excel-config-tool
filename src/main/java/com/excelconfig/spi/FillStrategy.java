package com.excelconfig.spi;

import org.apache.poi.ss.usermodel.Workbook;
import com.excelconfig.model.ExportConfig;

/**
 * 导出策略接口
 */
public interface FillStrategy {

    /**
     * 执行填充
     * @param workbook Excel Workbook
     * @param config 导出配置
     * @param context 填充上下文
     */
    void fill(Workbook workbook, ExportConfig config, FillContext context);

    /**
     * 支持的导出模式
     * @return 导出模式
     */
    FillMode getSupportedMode();
}
