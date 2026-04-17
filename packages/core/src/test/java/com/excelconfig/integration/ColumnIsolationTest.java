package com.excelconfig.integration;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.export.FillEngine;
import com.excelconfig.model.ExcelConfig;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 集成测试 - 列隔离与独立扩展
 *
 * 验证 A 列和 B 列在填充时互不影响，各自独立扩展
 */
public class ColumnIsolationTest {

    /**
     * 测试场景：
     * - A 列（姓名）：3 条数据
     * - B 列（年龄）：5 条数据
     *
     * 验证：A 列和 B 列各自独立扩展，互不干扰
     */
    @Test
    void testIndependentExpansion_DifferentColumnSizes() throws Exception {
        // 创建模板：每列只有 1 行数据空间
        byte[] template = createTwoColumnTemplate();

        // 配置：A 列和 B 列独立配置
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "names",
                  "header": { "match": "姓名" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "ages",
                  "header": { "match": "年龄" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列 3 条，B 列 5 条
        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三", "李四", "王五"));
        data.put("ages", Arrays.asList(20, 25, 30, 40, 50));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // A 列：1 表头 + 3 数据 = 4 行
            // B 列：1 表头 + 5 数据 = 6 行
            // 但由于列隔离，应该是 max(4, 6) = 6 行
            assertEquals(6, sheet.getPhysicalNumberOfRows());

            // 验证 A 列数据（姓名）
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("张三", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(0).getStringCellValue());
            // A 列第 4、5 行应该为空或不存在
            if (sheet.getRow(4) != null) {
                var cell = sheet.getRow(4).getCell(0);
                assertTrue(cell == null || cell.toString().isEmpty());
            }
            if (sheet.getRow(5) != null) {
                var cell = sheet.getRow(5).getCell(0);
                assertTrue(cell == null || cell.toString().isEmpty());
            }

            // 验证 B 列数据（年龄）
            assertEquals("年龄", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(20.0, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(25.0, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(30.0, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(40.0, sheet.getRow(4).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(50.0, sheet.getRow(5).getCell(1).getNumericCellValue(), 0.01);
        }
    }

    /**
     * 测试场景：
     * - A 列（姓名）：5 条数据
     * - B 列（年龄）：3 条数据
     *
     * 验证：A 列扩展不会影响 B 列已有的下方内容
     */
    @Test
    void testColumnA_ExpandDoesNotAffectColumnB() throws Exception {
        // 创建模板
        byte[] template = createTwoColumnTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "names",
                  "header": { "match": "姓名" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "ages",
                  "header": { "match": "年龄" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列 5 条，B 列 3 条
        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三", "李四", "王五", "赵六", "钱七"));
        data.put("ages", Arrays.asList(20, 25, 30));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 应该有 6 行（1 表头 + 5 数据，由 A 列决定）
            assertEquals(6, sheet.getPhysicalNumberOfRows());

            // 验证 A 列（姓名）- 5 条数据
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("张三", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals("赵六", sheet.getRow(4).getCell(0).getStringCellValue());
            assertEquals("钱七", sheet.getRow(5).getCell(0).getStringCellValue());

            // 验证 B 列（年龄）- 3 条数据
            assertEquals("年龄", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(20, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(25, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(30, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.01);
            // B 列第 4、5 行应该为空或不存在
            if (sheet.getRow(4) != null) {
                var cell = sheet.getRow(4).getCell(1);
                assertTrue(cell == null || cell.toString().isEmpty());
            }
            if (sheet.getRow(5) != null) {
                var cell = sheet.getRow(5).getCell(1);
                assertTrue(cell == null || cell.toString().isEmpty());
            }
        }
    }

    /**
     * 测试场景：
     * - A 列和 B 列都有原有数据
     * - A 列新增数据需要下移 B 列的内容
     *
     * 验证：列隔离机制确保互相不影响
     */
    @Test
    void testColumnIsolation_WithExistingData() throws Exception {
        // 创建模板：A 列和 B 列各有 2 行数据，下方有合计行
        byte[] template = createTemplateWithTotalRow();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "amounts",
                  "header": { "match": "金额" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列 3 条，B 列 2 条
        Map<String, Object> data = new HashMap<>();
        data.put("orderNos", Arrays.asList("ORD001", "ORD002", "ORD003"));
        data.put("amounts", Arrays.asList(100.00, 200.00));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 验证 A 列（订单号）- 3 条数据
            assertEquals("订单号", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("ORD001", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD002", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("ORD003", sheet.getRow(3).getCell(0).getStringCellValue());

            // 验证 B 列（金额）- 2 条数据
            assertEquals("金额", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(100.00, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(200.00, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.01);

            // 验证：每列各自填充，互不干扰
            // A 列填充 3 行，B 列填充 2 行
            // 总行数由实现决定，但关键是每列数据正确
        }
    }

    /**
     * 测试场景（核心列隔离）：
     * - 模板：A 列和 B 列各有 2 行原有数据（A3、B3 开始是原有数据）
     * - A 列配置 10 条新数据
     * - B 列配置 8 条新数据
     *
     * 预期结果：
     * - A1-A10: 新数据
     * - A11: 原来的 A3 值（被下移）
     * - B1-B8: 新数据
     * - B9: 原来的 B3 值（被下移）
     *
     * 验证：每列独立扩展，原有数据被正确下移
     */
    @Test
    void testColumnIsolation_CoreScenario() throws Exception {
        // 创建模板：每列有 2 行原有数据
        byte[] template = createTemplateWithMultipleRows();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "colA",
                  "header": { "match": "A 列" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "colB",
                  "header": { "match": "B 列" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列 10 条，B 列 8 条
        Map<String, Object> data = new HashMap<>();
        List<String> colAData = Arrays.asList("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10");
        List<String> colBData = Arrays.asList("B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8");
        data.put("colA", colAData);
        data.put("colB", colBData);

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 验证表头
            assertEquals("A 列", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("B 列", sheet.getRow(0).getCell(1).getStringCellValue());

            // 验证 A 列：10 条新数据
            for (int i = 0; i < 10; i++) {
                String expected = "A" + (i + 1);
                String actual = sheet.getRow(i + 1).getCell(0).getStringCellValue();
                assertEquals(expected, actual, "A 列第 " + (i + 1) + " 行应该是 " + expected);
            }

            // 验证 B 列：8 条新数据
            for (int i = 0; i < 8; i++) {
                String expected = "B" + (i + 1);
                String actual = sheet.getRow(i + 1).getCell(1).getStringCellValue();
                assertEquals(expected, actual, "B 列第 " + (i + 1) + " 行应该是 " + expected);
            }

            // 验证 A 列原有数据被下移
            // A 列填充 10 行（行 1-10），原有数据从行 11 开始
            // 原来行 1 的数据 "A-old-1" 应该在行 11
            // 原来行 2 的数据 "A-old-2" 应该在行 12
            // 原来行 3 的数据 "A-old-3" 应该在行 13
            assertNotNull(sheet.getRow(11), "A 列第 11 行应该存在");
            assertEquals("A-old-1", sheet.getRow(11).getCell(0).getStringCellValue());
            assertEquals("A-old-2", sheet.getRow(12).getCell(0).getStringCellValue());
            assertEquals("A-old-3", sheet.getRow(13).getCell(0).getStringCellValue());

            // 验证 B 列原有数据被下移
            // B 列填充 8 行（行 1-8），原有数据从行 9 开始
            // 原来行 1 的数据 "B-old-1" 应该在行 9
            // 原来行 2 的数据 "B-old-2" 应该在行 10
            // 原来行 3 的数据 "B-old-3" 应该在行 11
            assertNotNull(sheet.getRow(9), "B 列第 9 行应该存在");
            assertEquals("B-old-1", sheet.getRow(9).getCell(1).getStringCellValue());
            assertEquals("B-old-2", sheet.getRow(10).getCell(1).getStringCellValue());
            assertEquals("B-old-3", sheet.getRow(11).getCell(1).getStringCellValue());
        }
    }

    /**
     * 测试场景：
     * - 三列数据（A、B、C 列）
     * - 每列数据量不同：A 列 5 条，B 列 2 条，C 列 8 条
     *
     * 验证：多列情况下，各列独立扩展，互不影响
     */
    @Test
    void testThreeColumns_IndependentExpansion() throws Exception {
        byte[] template = createThreeColumnTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "names",
                  "header": { "match": "姓名" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "ages",
                  "header": { "match": "年龄" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "cities",
                  "header": { "match": "城市" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列 5 条，B 列 2 条，C 列 8 条
        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三", "李四", "王五", "赵六", "钱七"));
        data.put("ages", Arrays.asList(20, 25));
        data.put("cities", Arrays.asList("北京", "上海", "广州", "深圳", "杭州", "成都", "武汉", "西安"));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 验证 A 列（姓名）- 5 条数据
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            for (int i = 0; i < 5; i++) {
                assertNotNull(sheet.getRow(i + 1), "A 列第 " + (i + 1) + " 行应存在");
                assertNotNull(sheet.getRow(i + 1).getCell(0), "A 列第 " + (i + 1) + " 行单元格应存在");
            }
            assertEquals("钱七", sheet.getRow(5).getCell(0).getStringCellValue());

            // 验证 B 列（年龄）- 2 条数据
            assertEquals("年龄", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(20, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(25, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.01);
            // B 列第 3 行及以后应该为空
            if (sheet.getRow(3) != null && sheet.getRow(3).getCell(1) != null) {
                assertTrue(sheet.getRow(3).getCell(1).toString().isEmpty(), "B 列第 3 行应该为空");
            }

            // 验证 C 列（城市）- 8 条数据
            assertEquals("城市", sheet.getRow(0).getCell(2).getStringCellValue());
            for (int i = 0; i < 8; i++) {
                assertNotNull(sheet.getRow(i + 1), "C 列第 " + (i + 1) + " 行应存在");
                assertNotNull(sheet.getRow(i + 1).getCell(2), "C 列第 " + (i + 1) + " 行单元格应存在");
            }
            assertEquals("西安", sheet.getRow(8).getCell(2).getStringCellValue());
        }
    }

    /**
     * 测试场景：
     * - A 列和 B 列数据量相同
     *
     * 验证：两列数据量相同时，正常填充
     */
    @Test
    void testSameSizeColumns() throws Exception {
        byte[] template = createTwoColumnTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "names",
                  "header": { "match": "姓名" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "ages",
                  "header": { "match": "年龄" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列和 B 列都是 3 条
        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三", "李四", "王五"));
        data.put("ages", Arrays.asList(20, 25, 30));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 应该有 4 行（1 表头 + 3 数据）
            assertEquals(4, sheet.getPhysicalNumberOfRows());

            // 验证 A 列
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("张三", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(0).getStringCellValue());

            // 验证 B 列
            assertEquals("年龄", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(20, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(25, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(30, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.01);
        }
    }

    /**
     * 测试场景：
     * - A 列只有 1 条数据
     * - B 列有 10 条数据
     *
     * 验证：极端差异情况下，列隔离仍然有效
     */
    @Test
    void testExtremeDifference_ColumnSizes() throws Exception {
        byte[] template = createTwoColumnTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "names",
                  "header": { "match": "姓名" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "ages",
                  "header": { "match": "年龄" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列 1 条，B 列 10 条
        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三"));
        data.put("ages", Arrays.asList(20, 25, 30, 35, 40, 45, 50, 55, 60, 65));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 验证 A 列 - 只有 1 条数据
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("张三", sheet.getRow(1).getCell(0).getStringCellValue());
            // A 列第 2 行及以后应该为空
            if (sheet.getRow(2) != null && sheet.getRow(2).getCell(0) != null) {
                assertTrue(sheet.getRow(2).getCell(0).toString().isEmpty(), "A 列第 2 行应该为空");
            }

            // 验证 B 列 - 10 条数据
            assertEquals("年龄", sheet.getRow(0).getCell(1).getStringCellValue());
            for (int i = 0; i < 10; i++) {
                assertNotNull(sheet.getRow(i + 1), "B 列第 " + (i + 1) + " 行应存在");
                assertNotNull(sheet.getRow(i + 1).getCell(1), "B 列第 " + (i + 1) + " 行单元格应存在");
            }
            assertEquals(65, sheet.getRow(10).getCell(1).getNumericCellValue(), 0.01);
        }
    }

    /**
     * 测试场景：
     * - A 列为空列表
     * - B 列有数据
     *
     * 验证：空列表不影响其他列
     */
    @Test
    void testEmptyColumn_DoesNotAffectOthers() throws Exception {
        byte[] template = createTwoColumnTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "names",
                  "header": { "match": "姓名" },
                  "mode": "FILL_DOWN"
                },
                {
                  "key": "ages",
                  "header": { "match": "年龄" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：A 列为空，B 列 3 条
        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList());
        data.put("ages", Arrays.asList(20, 25, 30));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 验证 A 列 - 空列表，不应填充数据
            assertEquals("姓名", sheet.getRow(0).getCell(0).getStringCellValue());
            // A 列第 1 行应该为空或不存在

            // 验证 B 列 - 3 条数据正常填充
            assertEquals("年龄", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(20, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(25, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.01);
            assertEquals(30, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.01);
        }
    }

    // ===== 辅助方法 =====

    private byte[] createTwoColumnTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("姓名");
            headerRow.createCell(1).setCellValue("年龄");

            // 每列只有 1 行数据空间
            var dataRow = sheet.createRow(1);
            dataRow.createCell(0);
            dataRow.createCell(1);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createThreeColumnTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("姓名");
            headerRow.createCell(1).setCellValue("年龄");
            headerRow.createCell(2).setCellValue("城市");

            // 每列只有 1 行数据空间
            var dataRow = sheet.createRow(1);
            dataRow.createCell(0);
            dataRow.createCell(1);
            dataRow.createCell(2);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithMultipleRows() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("A 列");
            headerRow.createCell(1).setCellValue("B 列");

            // 2 行原有数据
            for (int i = 1; i <= 2; i++) {
                var row = sheet.createRow(i);
                row.createCell(0).setCellValue("A-old-" + i);
                row.createCell(1).setCellValue("B-old-" + i);
            }

            // 第 3 行开始也有原有数据（用于验证下移）
            var row3 = sheet.createRow(3);
            row3.createCell(0).setCellValue("A-old-3");
            row3.createCell(1).setCellValue("B-old-3");

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithTotalRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("订单号");
            headerRow.createCell(1).setCellValue("金额");
            headerRow.createCell(2).setCellValue("备注");

            // 2 行数据
            for (int i = 1; i <= 2; i++) {
                var row = sheet.createRow(i);
                row.createCell(0).setCellValue("ORD00" + i);
                row.createCell(1).setCellValue(100.00 * i);
                row.createCell(2).setCellValue("备注" + i);
            }

            // 合计行
            var totalRow = sheet.createRow(3);
            totalRow.createCell(2).setCellValue("合计");

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }
}
