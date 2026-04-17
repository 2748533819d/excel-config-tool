package com.excelconfig.integration;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.export.FillEngine;
import com.excelconfig.model.ExcelConfig;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 真实场景测试 - 覆盖多种业务场景
 *
 * 所有测试都创建真实的 Excel 文件并验证结果
 * 输出目录：/Users/huangzhenzhen/Documents/excel-test/
 */
public class RealFileColumnIsolationTest {

    private static final String OUTPUT_DIR = "/Users/huangzhenzhen/Documents/excel-test/";

    // ==================== 核心需求测试 ====================

    /**
     * 核心需求测试：
     * A 列配置 10 个数据向下填充，A1-A10 是新数据，A11 是原来模板的值
     * B 列配置 8 个数据向下填充，B1-B8 是新数据，B9 是原来模板的值
     */
    @Test
    void test01_CoreRequirement() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(3, "A", "B");
        saveToFile(template, "01-core-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 10));
        data.put("colB", generateData("B-", 8));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "01-core-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= 10; i++) {
                assertEquals("A-" + i, sheet.getRow(i).getCell(0).getStringCellValue());
            }
            assertEquals("A-old-1", sheet.getRow(11).getCell(0).getStringCellValue());

            for (int i = 1; i <= 8; i++) {
                assertEquals("B-" + i, sheet.getRow(i).getCell(1).getStringCellValue());
            }
            assertEquals("B-old-1", sheet.getRow(9).getCell(1).getStringCellValue());
        }

        System.out.println("✓ 核心需求测试通过");
    }

    // ==================== 数据量差异测试 ====================

    /**
     * 测试：A 列数据量远大于 B 列（100 vs 10）
     */
    @Test
    void test02_LargeDifference() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(5, "A", "B");
        saveToFile(template, "02-large-diff-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 100));
        data.put("colB", generateData("B-", 10));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "02-large-diff-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-100", sheet.getRow(100).getCell(0).getStringCellValue());
            assertEquals("B-10", sheet.getRow(10).getCell(1).getStringCellValue());
            assertEquals("B-old-1", sheet.getRow(11).getCell(1).getStringCellValue());
        }

        System.out.println("✓ 大数据量差异测试通过");
    }

    /**
     * 测试：A 列数据量远小于 B 列（5 vs 50）
     */
    @Test
    void test03_SmallVsLarge() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(3, "A", "B");
        saveToFile(template, "03-small-large-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 5));
        data.put("colB", generateData("B-", 50));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "03-small-large-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-5", sheet.getRow(5).getCell(0).getStringCellValue());
            assertEquals("B-50", sheet.getRow(50).getCell(1).getStringCellValue());
            assertEquals("A-old-1", sheet.getRow(6).getCell(0).getStringCellValue());
        }

        System.out.println("✓ 小 vs 大数据量测试通过");
    }

    // ==================== 多列测试 ====================

    /**
     * 测试：3 列不同数据量
     */
    @Test
    void test04_ThreeColumns() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(2, "A", "B", "C");
        saveToFile(template, "04-three-cols-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"},
                {"key": "colC", "header": {"match": "C"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 20));
        data.put("colB", generateData("B-", 15));
        data.put("colC", generateData("C-", 30));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "04-three-cols-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-20", sheet.getRow(20).getCell(0).getStringCellValue());
            assertEquals("B-15", sheet.getRow(15).getCell(1).getStringCellValue());
            assertEquals("C-30", sheet.getRow(30).getCell(2).getStringCellValue());
        }

        System.out.println("✓ 三列测试通过");
    }

    /**
     * 测试：5 列不同数据量
     */
    @Test
    void test05_FiveColumns() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(2, "A", "B", "C", "D", "E");
        saveToFile(template, "05-five-cols-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"},
                {"key": "colC", "header": {"match": "C"}, "mode": "FILL_DOWN"},
                {"key": "colD", "header": {"match": "D"}, "mode": "FILL_DOWN"},
                {"key": "colE", "header": {"match": "E"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 10));
        data.put("colB", generateData("B-", 20));
        data.put("colC", generateData("C-", 5));
        data.put("colD", generateData("D-", 50));
        data.put("colE", generateData("E-", 1));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "05-five-cols-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-10", sheet.getRow(10).getCell(0).getStringCellValue());
            assertEquals("B-20", sheet.getRow(20).getCell(1).getStringCellValue());
            assertEquals("C-5", sheet.getRow(5).getCell(2).getStringCellValue());
            assertEquals("D-50", sheet.getRow(50).getCell(3).getStringCellValue());
            assertEquals("E-1", sheet.getRow(1).getCell(4).getStringCellValue());
        }

        System.out.println("✓ 五列测试通过");
    }

    // ==================== 边界情况测试 ====================

    /**
     * 测试：空列表（不填充数据）
     */
    @Test
    void test06_EmptyList() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(3, "A", "B");
        saveToFile(template, "06-empty-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", Arrays.asList());
        data.put("colB", Arrays.asList());

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "06-empty-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("A-old-1", sheet.getRow(1).getCell(0).getStringCellValue());
        }

        System.out.println("✓ 空列表测试通过");
    }

    /**
     * 测试：只有一列有数据
     */
    @Test
    void test07_SingleColumnWithData() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(2, "A", "B", "C");
        saveToFile(template, "07-single-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"},
                {"key": "colC", "header": {"match": "C"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 10));
        data.put("colB", Arrays.asList());
        data.put("colC", Arrays.asList());

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "07-single-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-10", sheet.getRow(10).getCell(0).getStringCellValue());
            assertEquals("B-old-1", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("C-old-1", sheet.getRow(1).getCell(2).getStringCellValue());
        }

        System.out.println("✓ 单列有数据测试通过");
    }

    /**
     * 测试：只有 1 条数据
     */
    @Test
    void test08_SingleDataItem() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(5, "A", "B");
        saveToFile(template, "08-single-item-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", Arrays.asList("A-only"));
        data.put("colB", Arrays.asList("B-only"));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "08-single-item-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-only", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("B-only", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("A-old-1", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("B-old-1", sheet.getRow(2).getCell(1).getStringCellValue());
        }

        System.out.println("✓ 单条数据测试通过");
    }

    // ==================== 数据类型测试 ====================

    /**
     * 测试：数字类型数据
     */
    @Test
    void test09_NumericData() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplateWithNumericData();
        saveToFile(template, "09-numeric-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "amount", "header": {"match": "金额"}, "mode": "FILL_DOWN"},
                {"key": "count", "header": {"match": "数量"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("amount", Arrays.asList(100.5, 200.5, 300.5, 400.5, 500.5));
        data.put("count", Arrays.asList(10, 20, 30));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "09-numeric-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals(100.5, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.01);
            assertEquals(500.5, sheet.getRow(5).getCell(0).getNumericCellValue(), 0.01);
            assertEquals(30, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.01);
        }

        System.out.println("✓ 数字类型测试通过");
    }

    /**
     * 测试：混合类型数据（文本、数字、布尔）
     */
    @Test
    void test10_MixedDataTypes() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(2, "A", "B");
        saveToFile(template, "10-mixed-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("colA", Arrays.asList("文本", 123, true, "更多文本"));
        data.put("colB", Arrays.asList(99.99, "B 文本"));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "10-mixed-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("文本", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(123, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.01);
            assertEquals(99.99, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals("B 文本", sheet.getRow(2).getCell(1).getStringCellValue());
        }

        System.out.println("✓ 混合类型测试通过");
    }

    // ==================== 带有合计行测试 ====================

    /**
     * 测试：模板底部有合计行
     */
    @Test
    void test11_WithTotalRow() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplateWithTotalRow();
        saveToFile(template, "11-total-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "orderNo", "header": {"match": "订单号"}, "mode": "FILL_DOWN"},
                {"key": "amount", "header": {"match": "金额"}, "mode": "FILL_DOWN"},
                {"key": "remark", "header": {"match": "备注"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        Map<String, Object> data = new HashMap<>();
        data.put("orderNo", generateData("ORD-", 5));
        data.put("amount", Arrays.asList(100, 200, 300, 400, 500, 600, 700));
        data.put("remark", Arrays.asList("注 1", "注 2", "注 3", "注 4", "注 5", "注 6", "注 7"));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);
        saveToFile(result, "11-total-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            assertEquals("ORD-1", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD-5", sheet.getRow(5).getCell(0).getStringCellValue());
            assertEquals(700.0, sheet.getRow(7).getCell(1).getNumericCellValue(), 0.01);

            // 合计行应该被下移（备注列填充 7 行，原有数据 2 行，合计行在第 3 行，下移 7 行到第 10 行）
            boolean foundTotal = false;
            for (int i = 8; i <= sheet.getLastRowNum(); i++) {
                var row = sheet.getRow(i);
                if (row != null && row.getCell(2) != null) {
                    if ("合计".equals(row.getCell(2).getStringCellValue())) {
                        foundTotal = true;
                        System.out.println("合计行在第 " + i + " 行");
                        break;
                    }
                }
            }
            assertTrue(foundTotal, "应该找到合计行");
        }

        System.out.println("✓ 合计行测试通过");
    }

    // ==================== 性能相关测试 ====================

    /**
     * 测试：大数据量（1000 行）
     */
    @Test
    void test12_LargeDataVolume() throws Exception {
        ensureOutputDir();

        byte[] template = createTemplate(2, "A", "B", "C");
        saveToFile(template, "12-large-template.xlsx");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "colA", "header": {"match": "A"}, "mode": "FILL_DOWN"},
                {"key": "colB", "header": {"match": "B"}, "mode": "FILL_DOWN"},
                {"key": "colC", "header": {"match": "C"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        long startTime = System.currentTimeMillis();

        Map<String, Object> data = new HashMap<>();
        data.put("colA", generateData("A-", 1000));
        data.put("colB", generateData("B-", 500));
        data.put("colC", generateData("C-", 800));

        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        long duration = System.currentTimeMillis() - startTime;
        System.out.println("1000 行填充耗时：" + duration + "ms");
        saveToFile(result, "12-large-result.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("A-1000", sheet.getRow(1000).getCell(0).getStringCellValue());
            assertEquals("B-500", sheet.getRow(500).getCell(1).getStringCellValue());
            assertEquals("C-800", sheet.getRow(800).getCell(2).getStringCellValue());
        }

        assertTrue(duration < 5000, "应该在 5 秒内完成，实际：" + duration + "ms");
        System.out.println("✓ 大数据量测试通过（" + duration + "ms）");
    }

    // ==================== 辅助方法 ====================

    private void ensureOutputDir() {
        java.io.File dir = new java.io.File(OUTPUT_DIR);
        if (!dir.exists()) {
            dir.mkdirs();
        }
    }

    private byte[] createTemplate(int dataRows, String... headers) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            var headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            for (int i = 1; i <= dataRows; i++) {
                var row = sheet.createRow(i);
                for (int j = 0; j < headers.length; j++) {
                    row.createCell(j).setCellValue(headers[j] + "-old-" + i);
                }
            }

            java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithNumericData() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("金额");
            headerRow.createCell(1).setCellValue("数量");

            for (int i = 1; i <= 3; i++) {
                var row = sheet.createRow(i);
                row.createCell(0).setCellValue(1000.0 + i);
                row.createCell(1).setCellValue(i * 10);
            }

            java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithTotalRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("订单号");
            headerRow.createCell(1).setCellValue("金额");
            headerRow.createCell(2).setCellValue("备注");

            for (int i = 1; i <= 2; i++) {
                var row = sheet.createRow(i);
                row.createCell(0).setCellValue("ORD-OLD-" + i);
                row.createCell(1).setCellValue(100.0 * i);
                row.createCell(2).setCellValue("备注" + i);
            }

            var totalRow = sheet.createRow(3);
            totalRow.createCell(2).setCellValue("合计");

            java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private void saveToFile(byte[] data, String filename) throws IOException {
        String path = OUTPUT_DIR + filename;
        try (FileOutputStream fos = new FileOutputStream(path)) {
            fos.write(data);
        }
    }

    private java.util.List<String> generateData(String prefix, int count) {
        java.util.List<String> list = new java.util.ArrayList<>();
        for (int i = 1; i <= count; i++) {
            list.add(prefix + i);
        }
        return list;
    }
}
