package com.excelconfig;

import com.excelconfig.model.ExcelConfig;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 门面服务测试
 */
public class ExcelConfigServiceTest {

    @Test
    void testExtract() throws Exception {
        // 创建测试数据
        byte[] template = createTemplate();

        String configJson = """
            {
              "version": "1.0",
              "extractions": [
                {"key": "names", "header": {"match": "姓名"}, "mode": "DOWN"},
                {"key": "ages", "header": {"match": "年龄"}, "mode": "DOWN"}
              ]
            }
            """;

        ExcelConfigService service = new ExcelConfigService();
        Map<String, Object> result = service.extract(new ByteArrayInputStream(template), configJson);

        assertNotNull(result);
        assertTrue(result.containsKey("names"));
        assertTrue(result.containsKey("ages"));
        assertEquals(3, ((java.util.List<?>) result.get("names")).size());
        assertEquals(3, ((java.util.List<?>) result.get("ages")).size());

        System.out.println("✓ extract 测试通过");
    }

    @Test
    void testFill() throws Exception {
        // 创建测试模板
        byte[] template = createTemplateWithHeaders();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "names", "header": {"match": "姓名"}, "mode": "FILL_DOWN"},
                {"key": "ages", "header": {"match": "年龄"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三", "李四", "王五"));
        data.put("ages", Arrays.asList(20, 25, 30));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证结果
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("张三", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals(30, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.01);
        }

        System.out.println("✓ fill 测试通过");
    }

    @Test
    void testFill_ToOutputStream() throws Exception {
        byte[] template = createTemplateWithHeaders();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "names", "header": {"match": "姓名"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        data.put("names", Arrays.asList("张三", "李四"));

        ExcelConfigService service = new ExcelConfigService();
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        service.fill(new ByteArrayInputStream(template), data, configJson, output);

        // 验证输出
        assertTrue(output.size() > 0, "应该有输出数据");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("张三", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(0).getStringCellValue());
        }

        System.out.println("✓ fill ToOutputStream 测试通过");
    }

    @Test
    void testParseConfig() {
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {"key": "data", "header": {"match": "数据"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfigService service = new ExcelConfigService();
        ExcelConfig config = service.parseConfig(configJson);

        assertNotNull(config);
        assertEquals("1.0", config.getVersion());
        assertEquals(1, config.getExports().size());
        assertEquals("data", config.getExports().get(0).getKey());

        System.out.println("✓ parseConfig 测试通过");
    }

    @Test
    void testExtractAndFill_RoundTrip() throws Exception {
        // 创建初始数据
        byte[] template = createTemplate();

        String extractConfig = """
            {
              "version": "1.0",
              "extractions": [
                {"key": "names", "header": {"match": "姓名"}, "mode": "DOWN"},
                {"key": "ages", "header": {"match": "年龄"}, "mode": "DOWN"}
              ]
            }
            """;

        String fillConfig = """
            {
              "version": "1.0",
              "exports": [
                {"key": "names", "header": {"match": "姓名"}, "mode": "FILL_DOWN"},
                {"key": "ages", "header": {"match": "年龄"}, "mode": "FILL_DOWN"}
              ]
            }
            """;

        ExcelConfigService service = new ExcelConfigService();

        // 提取
        Map<String, Object> extracted = service.extract(new ByteArrayInputStream(template), extractConfig);

        // 修改数据
        @SuppressWarnings("unchecked")
        java.util.List<Object> names = (java.util.List<Object>) extracted.get("names");
        names.set(0, "新名字");

        // 填充
        byte[] result = service.fill(new ByteArrayInputStream(template), extracted, fillConfig);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("新名字", sheet.getRow(1).getCell(0).getStringCellValue());
        }

        System.out.println("✓ 提取填充往返测试通过");
    }

    // ===== 辅助方法 =====

    private byte[] createTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("姓名");
            headerRow.createCell(1).setCellValue("年龄");

            for (int i = 1; i <= 3; i++) {
                var row = sheet.createRow(i);
                row.createCell(0).setCellValue("姓名" + i);
                row.createCell(1).setCellValue(20 + i);
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createTemplateWithHeaders() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("姓名");
            headerRow.createCell(1).setCellValue("年龄");

            var dataRow = sheet.createRow(1);
            dataRow.createCell(0);
            dataRow.createCell(1);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }
}
