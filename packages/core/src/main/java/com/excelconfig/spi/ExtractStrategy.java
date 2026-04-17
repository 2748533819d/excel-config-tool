package com.excelconfig.spi;

import org.apache.poi.ss.usermodel.Sheet;
import com.excelconfig.model.ExtractConfig;

import java.util.List;

/**
 * 提取策略接口
 */
public interface ExtractStrategy {

    /**
     * 执行提取
     * @param sheet Excel Sheet
     * @param config 提取配置
     * @param context 提取上下文
     * @return 提取结果
     */
    List<Object> extract(Sheet sheet, ExtractConfig config, ExtractContext context);

    /**
     * 支持的提取模式
     * @return 提取模式
     */
    com.excelconfig.spi.ExtractMode getSupportedMode();
}
