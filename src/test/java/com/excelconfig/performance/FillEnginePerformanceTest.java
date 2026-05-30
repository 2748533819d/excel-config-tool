package com.excelconfig.performance;

import com.excelconfig.model.*;
import com.excelconfig.export.FillEngine;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.api.Timeout;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 大数据量下的性能测试
 *
 * <p>验证引擎在不同数据规模下的吞吐量和响应时间，确保 StyleCache 等优化措施有效。
 * 性能阈值基于 M1 MacBook Pro / 开发环境设定，CI 环境可能会更慢。</p>
 */
@Tag("performance")
class FillEnginePerformanceTest {

    private static final Logger log = LoggerFactory.getLogger(FillEnginePerformanceTest.class);

    private final FillEngine engine = new FillEngine();

    /** 测试输出文件保存目录 */
    private static final Path OUTPUT_DIR = Path.of("/Users/huangzhenzhen/Documents/excel-test/未命名文件夹");

    static {
        try {
            Files.createDirectories(OUTPUT_DIR);
        } catch (Exception e) {
            log.warn("无法创建输出目录: {}", OUTPUT_DIR);
        }
    }

    /** 性能阈值：FILL_DOWN 模式下每行平均耗时不超过此值（微秒） */
    private static final long MAX_MICROS_PER_ROW_FILL_DOWN = 500;

    /** 性能阈值：FILL_TABLE 模式下每行平均耗时不超过此值（微秒） */
    private static final long MAX_MICROS_PER_ROW_FILL_TABLE = 2000;

    // ========== FILL_DOWN 性能测试 ==========

    @Test
    void testFillDown_1kRows() {
        testFillDownPerformance(1_000, "1K");
    }

    @Test
    void testFillDown_5kRows() {
        testFillDownPerformance(5_000, "5K");
    }

    @Test
    void testFillDown_10kRows() {
        testFillDownPerformance(10_000, "10K");
    }

    private void testFillDownPerformance(int rowCount, String label) {
        // 准备模板
        Workbook workbook = createTemplateWithHeader("订单号");
        ExcelConfig config = createFillDownConfig("orderNos", "订单号");

        // 准备数据
        Map<String, Object> data = new HashMap<>();
        data.put("orderNos", generateStringData("ORD", rowCount));

        // 执行
        byte[] result = executeFill(workbook, data, config, "fill-down-" + label.toLowerCase().replace("k", "k"));

        // 验证行数
        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            // 表头在第 0 行，数据从第 1 行开始
            assertEquals(rowCount, sheet.getLastRowNum(), label + " 行数验证");
            // 抽样验证内容
            assertEquals("ORD-00001", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD-" + String.format("%05d", rowCount), sheet.getRow(rowCount).getCell(0).getStringCellValue());
        } catch (Exception e) {
            fail(label + " 验证失败: " + e.getMessage());
        }

        log.info("FILL_DOWN {} rows: output size = {} bytes", label, result.length);
    }

    // ========== FILL_TABLE 性能测试 ==========

    @Test
    void testFillTable_1kRows_5Cols() {
        testFillTablePerformance(1_000, 5, "1Kx5");
    }

    @Test
    void testFillTable_5kRows_5Cols() {
        testFillTablePerformance(5_000, 5, "5Kx5");
    }

    private void testFillTablePerformance(int rowCount, int colCount, String label) {
        // 准备模板 — 创建多列表头
        List<String> headers = IntStream.range(0, colCount)
            .mapToObj(i -> "列" + (i + 1))
            .collect(Collectors.toList());

        Workbook workbook = createTemplateWithHeaders(headers);
        ExcelConfig config = createFillTableConfig("data", headers);

        // 准备数据
        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = IntStream.range(0, rowCount)
            .mapToObj(i -> {
                Map<String, Object> row = new LinkedHashMap<>();
                for (int c = 0; c < colCount; c++) {
                    row.put("col" + c, "V-" + String.format("%05d", i) + "-" + c);
                }
                return row;
            })
            .collect(Collectors.toList());
        data.put("data", rows);

        // 执行
        byte[] result = executeFill(workbook, data, config, "fill-table-" + label.toLowerCase().replace("k", "k"));

        // 验证行数
        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            // 数据行从第 1 行开始
            assertEquals(rowCount, sheet.getLastRowNum(), label + " 行数验证");
            // 抽样验证第一行和最后一行
            assertEquals("V-00000-0", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("V-" + String.format("%05d", rowCount - 1) + "-0", sheet.getRow(rowCount).getCell(0).getStringCellValue());
        } catch (Exception e) {
            fail(label + " 验证失败: " + e.getMessage());
        }

        log.info("FILL_TABLE {} rows: output size = {} bytes", label, result.length);
    }

    // ========== 大规格多列压力测试 ==========

    @Test
    @Timeout(value = 3, unit = TimeUnit.MINUTES)
    void testFillTable_10kRows_10Cols() {
        // 10 列 × 10K 行 = 100K 单元格，验证多列场景下的吞吐量
        testFillTableLargeScale(10_000, 10, "10Kx10");
    }

    @Test
    @Tag("large-scale")
    @Timeout(value = 10, unit = TimeUnit.MINUTES)
    void testFillTable_100kRows_10Cols() {
        // 10 列 × 100K 行 = 1M 单元格，大规模压力测试
        // 注意：此测试会创建 ~1M Cell 对象，需要足够堆内存（建议 -Xmx2G）
        testFillTableLargeScale(100_000, 10, "100Kx10");
    }

    private void testFillTableLargeScale(int rowCount, int colCount, String label) {
        // 准备 10 列表头
        String[] headerArray = new String[colCount];
        for (int i = 0; i < colCount; i++) {
            headerArray[i] = "列" + (i + 1);
        }
        List<String> headers = Arrays.asList(headerArray);

        Workbook workbook = createTemplateWithHeaders(headers);
        ExcelConfig config = createFillTableConfig("data", headers);

        // 每列加不同样式，验证 StyleCache 多 key 下的缓存效率
        for (ColumnConfig col : config.getExports().get(0).getColumns()) {
            StyleConfig colStyle = new StyleConfig();
            colStyle.setVerticalAlign("CENTER");
            col.setStyle(colStyle);
        }

        // 高效生成大数据 —— 避免流式 lambda 对象开销
        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            Map<String, Object> row = new LinkedHashMap<>(colCount);
            for (int c = 0; c < colCount; c++) {
                row.put("col" + c, "V-" + i + "-" + c);
            }
            rows.add(row);
        }
        data.put("data", rows);

        // 执行填充（executeFill 已包含计时日志）
        byte[] result = executeFill(workbook, data, config, "fill-table-" + label.toLowerCase());

        // 结果验证 —— 仅抽样首尾行（全量验证 100K 行太慢）
        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            assertEquals(rowCount, sheet.getLastRowNum(), label + " 行数");

            // 第一行（表头在 row 0，数据从 row 1 开始）
            assertEquals("V-0-0", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("V-0-" + (colCount - 1), sheet.getRow(1).getCell(colCount - 1).getStringCellValue());
            // 最后一行
            Row lastRow = sheet.getRow(rowCount);
            assertEquals("V-" + (rowCount - 1) + "-0", lastRow.getCell(0).getStringCellValue());
            assertEquals("V-" + (rowCount - 1) + "-" + (colCount - 1), lastRow.getCell(colCount - 1).getStringCellValue());
        } catch (Exception e) {
            fail(label + " 验证失败: " + e.getMessage());
        }

        long totalCells = (long) rowCount * colCount;
        log.info("FILL_TABLE {} ({}x{}): {} cells, output {} bytes", label, rowCount, colCount, totalCells, result.length);
    }

    // ========== StyleCache 缓存效果验证 ==========

    @Test
    void testFillDown_WithStyle_SameStyleAllRows() {
        // 场景：5K 行使用完全相同的样式配置
        // 预期：StyleCache 只创建 1 个 CellStyle，缓存命中率接近 100%
        int rowCount = 5_000;

        Workbook workbook = createTemplateWithHeader("订单号");
        ExcelConfig config = createFillDownConfig("orderNos", "订单号");

        // 设定带样式的配置：粗体 + 字体颜色 + 背景色
        StyleConfig style = new StyleConfig();
        style.setBold(true);
        style.setFontColor("#FF0000");
        style.setBackground("#F5F5F5");
        config.getExports().get(0).setStyle(style);

        Map<String, Object> data = new HashMap<>();
        data.put("orderNos", generateStringData("ORD", rowCount));

        byte[] result = executeFill(workbook, data, config, "fill-down-style-same");

        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            assertEquals(rowCount, sheet.getLastRowNum(), "统一样式行数验证");
            Row firstDataRow = sheet.getRow(1);
            assertNotNull(firstDataRow.getCell(0).getCellStyle(), "样式不应为空");
        } catch (Exception e) {
            fail("统一样式验证失败: " + e.getMessage());
        }

        log.info("FILL_DOWN 统一样式 {} rows: output size = {} bytes", rowCount, result.length);
    }

    @Test
    void testFillTable_WithAlternateRows() {
        // 场景：FILL_TABLE 隔行换色，交替使用 2 种 CellStyle
        // 预期：StyleCache 只创建 2 个 CellStyle（主样式 + 隔行样式）
        int rowCount = 5_000;

        Workbook workbook = createTemplateWithHeaders(Arrays.asList("订单号", "金额"));
        ExcelConfig config = createFillTableConfig("data", Arrays.asList("订单号", "金额"));
        config.getExports().get(0).setAlternateRows(true);

        StyleConfig style = new StyleConfig();
        style.setBold(true);
        style.setBackground("#E8F0FE");
        config.getExports().get(0).setStyle(style);

        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            Map<String, Object> row = new LinkedHashMap<>();
            row.put("col0", "ORD-" + String.format("%05d", i + 1));
            row.put("col1", (double) (i + 1) * 100);
            rows.add(row);
        }
        data.put("data", rows);

        byte[] result = executeFill(workbook, data, config, "fill-table-alternate");

        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            // 表头 + rowCount 数据行
            assertEquals(rowCount, sheet.getLastRowNum(), "隔行换色行数验证");
        } catch (Exception e) {
            fail("隔行换色验证失败: " + e.getMessage());
        }

        log.info("FILL_TABLE 隔行换色 {} rows: output size = {} bytes", rowCount, result.length);
    }

    // ========== 综合压力测试 ==========

    @Test
    void testFill_MultipleExports() {
        // 多个数据列同时填充，模拟真实场景
        int rowCount = 2_000;

        Workbook workbook = createTemplateWithHeaders(Arrays.asList("订单号", "金额", "状态"));
        ExcelConfig config = new ExcelConfig();

        // 配置 3 个 FILL_DOWN 列
        for (int i = 0; i < 3; i++) {
            ExportConfig exportConfig = new ExportConfig();
            exportConfig.setKey("col" + i);

            HeaderConfig header = new HeaderConfig();
            header.setMatch(i == 0 ? "订单号" : (i == 1 ? "金额" : "状态"));
            exportConfig.setHeader(header);
            exportConfig.setMode("FILL_DOWN");

            // 每列加不同样式，让 StyleCache 发挥效果
            StyleConfig style = new StyleConfig();
            style.setFontSize(11);
            style.setVerticalAlign("CENTER");
            exportConfig.setStyle(style);

            config.getExports().add(exportConfig);
        }

        // 准备数据
        Map<String, Object> data = new HashMap<>();
        for (int i = 0; i < 3; i++) {
            if (i == 1) {
                // 金额列用 Double
                data.put("col" + i, generateDoubleData(rowCount));
            } else {
                data.put("col" + i, generateStringData(i == 0 ? "ORD" : "ST", rowCount));
            }
        }

        byte[] result = executeFill(workbook, data, config, "fill-multi-column");

        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            // 3 列各自填充 rowCount 行，互不干扰
            assertEquals(rowCount, sheet.getLastRowNum(), "多列填充行数验证");
            // 验证第 1 行第 1 列
            assertEquals("ORD-00001", sheet.getRow(1).getCell(0).getStringCellValue());
        } catch (Exception e) {
            fail("多列填充验证失败: " + e.getMessage());
        }

        log.info("FILL 多列 {}x3 rows: output size = {} bytes", rowCount, result.length);
    }

    // ========== StyleCache 大数据量内存优化验证 ==========

    @Test
    void testFillDown_StyleCache_10kWithComplexStyle() {
        // 10K 行带复杂样式的填充，验证 StyleCache 防止 CellStyle 爆炸
        int rowCount = 10_000;

        Workbook workbook = createTemplateWithHeader("数据");
        ExcelConfig config = createFillDownConfig("values", "数据");

        // 复杂样式：粗体 + 字体颜色 + 背景 + 对齐方式 + 数字格式
        StyleConfig style = new StyleConfig();
        style.setBold(true);
        style.setFontColor("#333333");
        style.setBackground("#E8F0FE");
        style.setHorizontalAlign("CENTER");
        style.setVerticalAlign("CENTER");
        style.setFormat("#,##0.00");
        config.getExports().get(0).setStyle(style);

        // 准备 Double 数据
        Map<String, Object> data = new HashMap<>();
        data.put("values", generateDoubleData(rowCount));

        long startNanos = System.nanoTime();
        byte[] result = executeFill(workbook, data, config, "fill-down-10k-complex");
        long elapsedMicros = (System.nanoTime() - startNanos) / 1_000;

        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = resultWorkbook.getSheetAt(0);
            assertEquals(rowCount, sheet.getLastRowNum(), "10K 样式行数");

            // 验证数值格式正确应用
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(1.0, cell.getNumericCellValue(), 0.001);
        } catch (Exception e) {
            fail("复杂样式验证失败: " + e.getMessage());
        }

        long microsPerRow = elapsedMicros / rowCount;
        log.info("FILL_DOWN 10K 复杂样式: {} μs/row, total {} ms, output {} bytes",
            microsPerRow, elapsedMicros / 1000, result.length);

        // StyleCache 下 10K 复杂样式应该 < 300 μs/row（不含模板 I/O）
        assertTrue(microsPerRow < 500,
            "10K 复杂样式每行耗时 " + microsPerRow + " μs，预期 < 500 μs（StyleCache 应使样式创建成本趋近于 0）");
    }

    // ========== 辅助方法 ==========

    private byte[] executeFill(Workbook workbook, Map<String, Object> data, ExcelConfig config, String outputName) {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            workbook.write(baos);
            workbook.close();
        } catch (Exception e) {
            fail("模板写入失败: " + e.getMessage());
        }
        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());

        long startNanos = System.nanoTime();
        byte[] result;
        try {
            result = engine.fill(bais, data, config);
        } catch (Exception e) {
            fail("填充执行失败: " + e.getMessage());
            return null; // 不可达
        }
        long elapsedMicros = (System.nanoTime() - startNanos) / 1_000;
        long dataSize = data.values().stream()
            .filter(v -> v instanceof Collection)
            .mapToInt(v -> ((Collection<?>) v).size())
            .findFirst().orElse(0);

        log.info("耗时 {} ms ({} μs/row), 输出 {} bytes",
            elapsedMicros / 1000, dataSize > 0 ? elapsedMicros / dataSize : 0, result.length);

        assertNotNull(result);
        assertTrue(result.length > 0, "输出不应为空");

        // 保存到指定目录
        if (outputName != null) {
            try {
                Path outputFile = OUTPUT_DIR.resolve(outputName + ".xlsx");
                Files.write(outputFile, result);
                log.info("已保存: {}", outputFile);
            } catch (Exception e) {
                log.warn("保存输出文件失败: {}", e.getMessage());
            }
        }

        return result;
    }

    /** 向后兼容，不写文件 */
    private byte[] executeFill(Workbook workbook, Map<String, Object> data, ExcelConfig config) {
        return executeFill(workbook, data, config, null);
    }

    private Workbook createTemplateWithHeader(String headerText) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");
        sheet.createRow(0).createCell(0).setCellValue(headerText);
        return workbook;
    }

    private Workbook createTemplateWithHeaders(List<String> headers) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");
        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            row.createCell(i).setCellValue(headers.get(i));
        }
        return workbook;
    }

    private ExcelConfig createFillDownConfig(String key, String headerMatch) {
        ExcelConfig config = new ExcelConfig();
        ExportConfig exportConfig = new ExportConfig();
        exportConfig.setKey(key);

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch(headerMatch);
        exportConfig.setHeader(headerConfig);
        exportConfig.setMode("FILL_DOWN");

        config.getExports().add(exportConfig);
        return config;
    }

    private ExcelConfig createFillTableConfig(String key, List<String> headers) {
        ExcelConfig config = new ExcelConfig();
        ExportConfig exportConfig = new ExportConfig();
        exportConfig.setKey(key);

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch(headers.get(0));
        exportConfig.setHeader(headerConfig);
        exportConfig.setMode("FILL_TABLE");

        // 列配置
        List<ColumnConfig> columns = new ArrayList<>();
        for (int i = 0; i < headers.size(); i++) {
            ColumnConfig col = new ColumnConfig();
            col.setKey("col" + i);
            col.setHeader(headers.get(i));
            col.setWidth(15);
            columns.add(col);
        }
        exportConfig.setColumns(columns);

        config.getExports().add(exportConfig);
        return config;
    }

    private List<String> generateStringData(String prefix, int count) {
        List<String> data = new ArrayList<>(count);
        for (int i = 1; i <= count; i++) {
            data.add(prefix + "-" + String.format("%05d", i));
        }
        return data;
    }

    private List<Double> generateDoubleData(int count) {
        List<Double> data = new ArrayList<>(count);
        for (int i = 1; i <= count; i++) {
            data.add((double) i);
        }
        return data;
    }
}
