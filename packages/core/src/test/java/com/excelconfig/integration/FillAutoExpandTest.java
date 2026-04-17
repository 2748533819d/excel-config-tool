package com.excelconfig.integration;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.export.FillEngine;
import com.excelconfig.extract.ExtractEngine;
import com.excelconfig.model.ExcelConfig;
import com.excelconfig.model.ExportConfig;
import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.HeaderConfig;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 集成测试 - 测试填充自动扩展功能
 */
public class FillAutoExpandTest {

    @Test
    void testFillDown_AutoExpand() throws Exception {
        // 创建模板：只有 1 行数据空间
        byte[] template = createTemplateWithOneRow();

        // 配置
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：5 条记录，远超模板空间
        Map<String, Object> data = new HashMap<>();
        data.put("orderNos", Arrays.asList("ORD001", "ORD002", "ORD003", "ORD004", "ORD005"));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果：应该自动扩展到 5 行
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 应该有 6 行（1 行表头 + 5 行数据）
            assertEquals(6, sheet.getPhysicalNumberOfRows());

            // 验证数据
            assertEquals("ORD001", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD002", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("ORD003", sheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals("ORD004", sheet.getRow(4).getCell(0).getStringCellValue());
            assertEquals("ORD005", sheet.getRow(5).getCell(0).getStringCellValue());
        }
    }

    @Test
    void testFillTable_AutoExpand() throws Exception {
        // 创建模板：只有表头
        byte[] template = createHeaderOnlyTemplate();

        // 配置
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "orders",
                  "header": { "match": "订单号" },
                  "mode": "FILL_TABLE",
                  "columns": [
                    {"key": "orderNo", "header": "订单号", "width": 15},
                    {"key": "amount", "header": "金额", "width": 12}
                  ]
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：3 条记录
        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> orders = Arrays.asList(
            Map.of("orderNo", "ORD001", "amount", 100.00),
            Map.of("orderNo", "ORD002", "amount", 200.00),
            Map.of("orderNo", "ORD003", "amount", 300.00)
        );
        data.put("orders", orders);

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果：应该自动创建 3 行数据
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 应该有 4 行（1 行表头 + 3 行数据）
            assertEquals(4, sheet.getPhysicalNumberOfRows());

            // 验证表头
            assertEquals("订单号", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("金额", sheet.getRow(0).getCell(1).getStringCellValue());

            // 验证数据
            assertEquals("ORD001", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD002", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("ORD003", sheet.getRow(3).getCell(0).getStringCellValue());
        }
    }

    @Test
    void testFill_WithExistingDataBelow() throws Exception {
        // 创建模板：表头 + 1 行数据 + 下方有其他内容
        byte[] template = createTemplateWithDataBelow();

        // 配置
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = new JsonConfigParser().parse(configJson);

        // 数据：3 条记录，需要下移下方内容
        Map<String, Object> data = new HashMap<>();
        data.put("orderNos", Arrays.asList("ORD001", "ORD002", "ORD003"));

        // 执行填充
        FillEngine engine = new FillEngine();
        byte[] result = engine.fill(new ByteArrayInputStream(template), data, config);

        // 验证结果：下方内容应该被下移
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 验证数据
            assertEquals("ORD001", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD002", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("ORD003", sheet.getRow(3).getCell(0).getStringCellValue());

            // 下方的合计行应该被下移到第 5 行（原有行 2 + 下移 3 行 = 行 5）
            assertEquals("合计", sheet.getRow(5).getCell(0).getStringCellValue());
        }
    }

    @Test
    void testExtractAndFill_RoundTrip() throws Exception {
        // 创建初始数据
        byte[] template = createTemplateWithData();

        // 提取配置
        String extractJson = """
            {
              "version": "1.0",
              "extractions": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "DOWN"
                }
              ]
            }
            """;

        ExcelConfig extractConfig = new JsonConfigParser().parse(extractJson);

        // 提取数据
        ExtractEngine extractEngine = new ExtractEngine();
        Map<String, Object> extracted = extractEngine.extract(
            new ByteArrayInputStream(template),
            extractConfig
        );

        // 验证提取结果
        @SuppressWarnings("unchecked")
        List<Object> orderNos = (List<Object>) extracted.get("orderNos");
        assertEquals(3, orderNos.size());
        assertEquals("ORD001", orderNos.get(0));
        assertEquals("ORD002", orderNos.get(1));
        assertEquals("ORD003", orderNos.get(2));

        // 修改数据
        Map<String, Object> newData = new HashMap<>();
        newData.put("orderNos", Arrays.asList("NEW001", "NEW002", "NEW003", "NEW004", "NEW005"));

        // 填充配置
        String fillJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig fillConfig = new JsonConfigParser().parse(fillJson);

        // 执行填充
        FillEngine fillEngine = new FillEngine();
        byte[] result = fillEngine.fill(
            new ByteArrayInputStream(template),
            newData,
            fillConfig
        );

        // 验证填充结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);

            // 应该有 6 行（1 表头 + 5 数据）
            // 注：实际行数可能更多，因为 FILL_DOWN 会覆盖原有数据
            assertTrue(sheet.getPhysicalNumberOfRows() >= 6, "应该至少有 6 行");

            // 验证新数据
            assertEquals("NEW001", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("NEW005", sheet.getRow(5).getCell(0).getStringCellValue());
        }
    }

    // ===== 辅助方法 =====

    private byte[] createTemplateWithOneRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("订单号");

            // 只有 1 行数据空间
            var dataRow = sheet.createRow(1);
            dataRow.createCell(0);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createHeaderOnlyTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 只有表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("订单号");

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithDataBelow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("订单号");

            // 1 行数据
            var dataRow = sheet.createRow(1);
            dataRow.createCell(0);

            // 下方有合计行
            var totalRow = sheet.createRow(2);
            totalRow.createCell(0).setCellValue("合计");

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithData() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("订单号");

            // 3 行数据
            for (int i = 0; i < 3; i++) {
                var row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("ORD00" + (i + 1));
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }
}
