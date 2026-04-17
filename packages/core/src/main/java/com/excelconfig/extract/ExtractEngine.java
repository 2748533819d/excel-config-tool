package com.excelconfig.extract;

import com.excelconfig.locator.HeaderLocator;
import com.excelconfig.locator.HeaderPosition;
import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.ExcelConfig;
import com.excelconfig.model.HeaderConfig;
import com.excelconfig.model.PositionConfig;
import com.excelconfig.spi.ExtractContext;
import com.excelconfig.spi.ExtractStrategy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.InputStream;
import java.util.*;

/**
 * 提取引擎
 *
 * 核心功能：
 * 1. 表头自动定位
 * 2. 数据提取
 * 3. 边界检测
 */
public class ExtractEngine {

    private final HeaderLocator headerLocator;
    private final Map<com.excelconfig.spi.ExtractMode, ExtractStrategy> strategies;

    public ExtractEngine() {
        this.headerLocator = new HeaderLocator();
        this.strategies = new EnumMap<>(com.excelconfig.spi.ExtractMode.class);
        registerBuiltInStrategies();
    }

    /**
     * 注册内置策略
     */
    private void registerBuiltInStrategies() {
        // 注册基础策略
        registerStrategy(new SingleExtractStrategy());
        registerStrategy(new DownExtractStrategy());
        registerStrategy(new RightExtractStrategy());
        registerStrategy(new BlockExtractStrategy());
        registerStrategy(new UntilEmptyExtractStrategy());
        // 注册 SAX 流式策略
        registerStrategy(new SaxDownExtractStrategy());
    }

    /**
     * 注册提取策略
     */
    public void registerStrategy(ExtractStrategy strategy) {
        strategies.put(strategy.getSupportedMode(), strategy);
    }

    /**
     * 执行提取
     *
     * @param input Excel 文件输入流
     * @param config 配置
     * @return 提取结果 Map<String, Object>
     */
    public Map<String, Object> extract(InputStream input, ExcelConfig config) {
        try {
            Workbook workbook = WorkbookFactory.create(input);
            Sheet sheet = workbook.getSheetAt(0);

            Map<String, Object> result = new HashMap<>();

            for (ExtractConfig extractConfig : config.getExtractions()) {
                List<Object> data = extract(sheet, extractConfig);
                result.put(extractConfig.getKey(), data);
            }

            workbook.close();
            return result;

        } catch (Exception e) {
            throw new ExtractException("提取失败：" + e.getMessage(), e);
        }
    }

    /**
     * 提取单个配置
     */
    public List<Object> extract(Sheet sheet, ExtractConfig config) {
        try {
            // 1. 定位表头
            HeaderPosition headerPos = locateHeader(sheet, config);

            // 2. 获取提取策略
            com.excelconfig.spi.ExtractMode mode = com.excelconfig.spi.ExtractMode.fromString(config.getMode());
            ExtractStrategy strategy = strategies.get(mode);

            if (strategy == null) {
                throw new ExtractException("不支持的提取模式：" + config.getMode());
            }

            // 3. 创建上下文并执行提取
            ExtractContext context = new ExtractContext(
                config,
                headerPos.getRow() + 1,  // 从表头下方开始
                headerPos.getColumn()
            );

            return strategy.extract(sheet, config, context);

        } catch (Exception e) {
            throw new ExtractException("提取失败 [" + config.getKey() + "]: " + e.getMessage(), e);
        }
    }

    /**
     * 定位表头
     */
    private HeaderPosition locateHeader(Sheet sheet, ExtractConfig config) {
        // 优先使用表头匹配
        if (config.getHeader() != null && config.getHeader().getMatch() != null) {
            return headerLocator.locate(sheet, config.getHeader());
        }

        // 使用固定位置
        if (config.getPosition() != null && config.getPosition().getCellRef() != null) {
            PositionConfig pos = config.getPosition();
            CellReference ref = new CellReference(pos.getCellRef());
            return new HeaderPosition(ref.getRow(), ref.getCol());
        }

        throw new ExtractException("必须配置 header 或 position");
    }
}
