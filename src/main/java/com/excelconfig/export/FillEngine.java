package com.excelconfig.export;

import com.excelconfig.locator.HeaderLocator;
import com.excelconfig.locator.HeaderPosition;
import com.excelconfig.model.ExportConfig;
import com.excelconfig.model.ExcelConfig;
import com.excelconfig.model.PositionConfig;
import com.excelconfig.extract.CellReference;
import com.excelconfig.spi.FillContext;
import com.excelconfig.spi.FillStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
    private Workbook lookupWorkbook;

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

        Workbook workbook = null;
        Workbook beforeTemplate = null;
        try {
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook) WorkbookFactory.create(template);

            if (Boolean.TRUE.equals(config.getStreaming())) {
                // 流式写入：创建新 SXSSFWorkbook，复制模板内容（不包装，避免 SXSSF 无法覆写模板行的限制）
                int windowSize = config.getStreamingRowWindowSize() != null
                    ? config.getStreamingRowWindowSize() : 100;
                workbook = createStreamingWorkbook(xssfWorkbook, windowSize);
                // 保留模板 XSSFWorkbook 用于表头定位（SXSSF 的内部 XSSFWorkbook 是空的）
                this.lookupWorkbook = xssfWorkbook;
                beforeTemplate = xssfWorkbook;
                log.info("SXSSF 流式写入已启用，行缓存窗口 = {}", windowSize);
            } else {
                workbook = xssfWorkbook;
                this.lookupWorkbook = xssfWorkbook;
            }

            Workbook finalWorkbook = this.lookupWorkbook;

            // 按行号从下往上处理，避免覆盖
            List<ExportConfig> sortedExports = new ArrayList<>(config.getExports());
            sortedExports.sort((a, b) -> {
                int rowA = getStartRow(finalWorkbook, a);
                int rowB = getStartRow(finalWorkbook, b);
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
        } finally {
            this.lookupWorkbook = null;
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception ignored) {
                }
            }
            if (beforeTemplate != null && beforeTemplate != workbook) {
                try {
                    beforeTemplate.close();
                } catch (Exception ignored) {
                }
            }
        }
    }

    /**
     * 创建流式写入的 SXSSFWorkbook，将模板内容复制到全新的 SXSSF 工作簿中
     * （不直接包装 XSSFWorkbook，因为 SXSSF 不允许覆写包装中的已有行）
     */
    private static SXSSFWorkbook createStreamingWorkbook(XSSFWorkbook template, int windowSize) {
        SXSSFWorkbook swb = new SXSSFWorkbook(windowSize);
        for (int s = 0; s < template.getNumberOfSheets(); s++) {
            org.apache.poi.xssf.usermodel.XSSFSheet srcSheet = template.getSheetAt(s);
            SXSSFSheet destSheet = swb.createSheet(srcSheet.getSheetName());

            // 复制所有已有行（模板行，通常只有表头几行）
            for (int r = 0; r <= srcSheet.getLastRowNum(); r++) {
                Row srcRow = srcSheet.getRow(r);
                if (srcRow == null) continue;

                Row destRow = destSheet.createRow(r);
                for (Cell srcCell : srcRow) {
                    Cell destCell = destRow.createCell(srcCell.getColumnIndex());
                    copyCellValue(srcCell, destCell);
                    // 跨 workbook 克隆样式（SXSSF 的内部 XSSFWorkbook 是全新的，不能直接复用源 CellStyle）
                    CellStyle srcStyle = srcCell.getCellStyle();
                    if (srcStyle != null) {
                        CellStyle clonedStyle = com.excelconfig.util.StyleCache.cloneCellStyle(srcStyle, swb);
                        destCell.setCellStyle(clonedStyle);
                    }
                }
            }

            // 复制合并区域
            for (CellRangeAddress region : srcSheet.getMergedRegions()) {
                destSheet.addMergedRegion(region);
            }
        }
        return swb;
    }

    /**
     * 复制单元格值（不包含样式）
     */
    private static void copyCellValue(Cell src, Cell dest) {
        switch (src.getCellType()) {
            case STRING -> dest.setCellValue(src.getStringCellValue());
            case NUMERIC -> dest.setCellValue(src.getNumericCellValue());
            case BOOLEAN -> dest.setCellValue(src.getBooleanCellValue());
            case FORMULA -> dest.setCellFormula(src.getCellFormula());
            case BLANK -> dest.setBlank();
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
     * 定位表头（使用 lookupWorkbook，SXSSF 模式下与写入 Workbook 不同）
     */
    private HeaderPosition locateHeader(Workbook workbook, ExportConfig config) {
        Sheet sheet = (lookupWorkbook != null ? lookupWorkbook : workbook).getSheetAt(0);

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
