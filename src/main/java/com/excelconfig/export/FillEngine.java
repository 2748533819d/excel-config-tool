package com.excelconfig.export;

import com.excelconfig.locator.HeaderLocator;
import com.excelconfig.locator.HeaderPosition;
import com.excelconfig.model.ExportConfig;
import com.excelconfig.model.ExcelConfig;
import com.excelconfig.model.PositionConfig;
import com.excelconfig.extract.CellReference;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.InputStream;
import java.util.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 导出/填充引擎
 *
 * 核心功能：
 * 1. 表头自动定位
 * 2. 数据填充
 * 3. 动态扩展（空间不足时自动下移下方内容）
 */
public class FillEngine {

    private static final Logger log = LoggerFactory.getLogger(FillEngine.class);

    private final HeaderLocator headerLocator;
    private final Map<com.excelconfig.spi.FillMode, FillStrategy> strategies;

    public FillEngine() {
        this.headerLocator = new HeaderLocator();
        this.strategies = new EnumMap<>(com.excelconfig.spi.FillMode.class);
        registerBuiltInStrategies();
    }

    /**
     * 注册内置策略
     */
    private void registerBuiltInStrategies() {
        registerStrategy(new FillCellStrategy());
        registerStrategy(new FillDownStrategy());
        registerStrategy(new FillTableStrategy());
    }

    /**
     * 注册填充策略
     *
     * @param strategy 填充策略实例
     */
    private void registerStrategy(FillStrategy strategy) {
        strategies.put(strategy.getSupportedMode(), strategy);
    }

    /**
     * 执行填充
     *
     * @param template Excel 模板输入流
     * @param data 数据
     * @param config 配置
     * @return 填充后的 Excel 文件字节数组
     * @throws FillException 填充失败时抛出
     */
    public byte[] fill(InputStream template, Map<String, Object> data, ExcelConfig config) {
        log.info("开始填充数据：{} 个配置项", config.getExports().size());
        long startTime = System.currentTimeMillis();

        try (Workbook workbook = WorkbookFactory.create(template)) {

            // 按行号从下往上处理，避免覆盖
            List<ExportConfig> sortedExports = new ArrayList<>(config.getExports());
            sortedExports.sort((a, b) -> {
                int rowA = getStartRow(workbook, a);
                int rowB = getStartRow(workbook, b);
                return Integer.compare(rowB, rowA);  // 从下往上
            });

            for (ExportConfig exportConfig : sortedExports) {
                log.debug("填充配置 [{}] 模式={}", exportConfig.getKey(), exportConfig.getMode());
                fill(workbook, data, exportConfig);
            }

            java.io.ByteArrayOutputStream output = new java.io.ByteArrayOutputStream();
            workbook.write(output);
            log.info("填充完成，耗时 {}ms，输出 {} 字节", System.currentTimeMillis() - startTime, output.size());
            return output.toByteArray();

        } catch (Exception e) {
            log.error("填充失败", e);
            throw new FillException("填充失败：" + e.getMessage(), e);
        }
    }

    /**
     * 填充单个配置
     */
    void fill(Workbook workbook, Map<String, Object> data, ExportConfig config) {
        try {
            // 1. 定位表头
            HeaderPosition headerPos = locateHeader(workbook, config);

            // 2. 获取填充策略
            com.excelconfig.spi.FillMode mode = com.excelconfig.spi.FillMode.fromString(config.getMode());
            FillStrategy strategy = strategies.get(mode);

            if (strategy == null) {
                throw new FillException("不支持的填充模式：" + config.getMode());
            }

            // 3. 创建上下文并执行填充
            FillContext context = new FillContext(
                config,
                data,
                headerPos.getRow() + 1,  // 从表头下方开始
                headerPos.getColumn()
            );

            strategy.fill(workbook, config, context);

        } catch (Exception e) {
            throw new FillException("填充失败 [" + config.getKey() + "]: " + e.getMessage(), e);
        }
    }

    /**
     * 定位表头
     */
    private HeaderPosition locateHeader(Workbook workbook, ExportConfig config) {
        Sheet sheet = workbook.getSheetAt(0);

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

        throw new FillException("必须配置 header 或 position");
    }

    /**
     * 获取配置起始行（用于排序）
     */
    private int getStartRow(Workbook workbook, ExportConfig config) {
        try {
            HeaderPosition pos = locateHeader(workbook, config);
            return pos.getRow();
        } catch (Exception e) {
            log.warn("定位导出配置 [{}] 的表头失败，使用默认行号 0: {}", config.getKey(), e.getMessage());
            return 0;
        }
    }
}
