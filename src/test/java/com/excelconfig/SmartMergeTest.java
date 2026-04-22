package com.excelconfig;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 智能合并单元格功能测试
 *
 * 测试按数据值自动合并相同值的单元格
 */
public class SmartMergeTest {

    @Test
    void testSmartMerge_BasicSameValues() throws Exception {
        // 创建测试模板
        byte[] template = createSimpleTemplate();

        // 配置：部门字段启用智能合并
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "departments",
                  "header": {"match": "部门"},
                  "mode": "FILL_DOWN",
                  "merge": {
                    "enabled": true
                  }
                }
              ]
            }
            """;

        // 数据：3 个技术部 + 2 个销售部
        Map<String, Object> data = new HashMap<>();
        data.put("departments", Arrays.asList("技术部", "技术部", "技术部", "销售部", "销售部"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 应该有 2 个合并区域
            assertEquals(2, sheet.getNumMergedRegions(), "应该有 2 个合并区域");

            // 第一个合并区域：A1-A3（技术部）
            CellRangeAddress merged1 = sheet.getMergedRegion(0);
            assertEquals(1, merged1.getFirstRow(), "第一个合并区域起始行应为 1");
            assertEquals(3, merged1.getLastRow(), "第一个合并区域结束行应为 3");
            assertEquals(0, merged1.getFirstColumn(), "起始列应为 0");
            assertEquals(0, merged1.getLastColumn(), "结束列应为 0");

            // 第二个合并区域：A4-A5（销售部）
            CellRangeAddress merged2 = sheet.getMergedRegion(1);
            assertEquals(4, merged2.getFirstRow(), "第二个合并区域起始行应为 4");
            assertEquals(5, merged2.getLastRow(), "第二个合并区域结束行应为 5");

            // 验证数据
            assertEquals("技术部", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("销售部", sheet.getRow(4).getCell(0).getStringCellValue());

            // 验证合并区域内的单元格为空
            assertTrue(sheet.getRow(2).getCell(0).getCellType() == CellType.BLANK);
            assertTrue(sheet.getRow(3).getCell(0).getCellType() == CellType.BLANK);
            assertTrue(sheet.getRow(5).getCell(0).getCellType() == CellType.BLANK);
        }

        System.out.println("✓ 智能合并基础测试通过");
    }

    @Test
    void testSmartMerge_NoMergeForSingleValue() throws Exception {
        // 创建测试模板
        byte[] template = createSimpleTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "departments",
                  "header": {"match": "部门"},
                  "mode": "FILL_DOWN",
                  "merge": {"enabled": true}
                }
              ]
            }
            """;

        // 数据：每个部门都不同，不应该合并
        Map<String, Object> data = new HashMap<>();
        data.put("departments", Arrays.asList("技术部", "销售部", "人事部"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 没有合并区域
            assertEquals(0, sheet.getNumMergedRegions(), "不应该有合并区域");

            // 验证数据都在
            assertEquals("技术部", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("销售部", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("人事部", sheet.getRow(3).getCell(0).getStringCellValue());
        }

        System.out.println("✓ 单个值不合并测试通过");
    }

    @Test
    void testSmartMerge_MultipleGroups() throws Exception {
        // 创建测试模板
        byte[] template = createSimpleTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "groups",
                  "header": {"match": "分组"},
                  "mode": "FILL_DOWN",
                  "merge": {"enabled": true}
                }
              ]
            }
            """;

        // 数据：A-A-B-B-B-C-C-C-C
        Map<String, Object> data = new HashMap<>();
        data.put("groups", Arrays.asList("A", "A", "B", "B", "B", "C", "C", "C", "C"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 3 个合并区域
            assertEquals(3, sheet.getNumMergedRegions());

            // A: R1-R2
            CellRangeAddress merged1 = sheet.getMergedRegion(0);
            assertEquals(1, merged1.getFirstRow());
            assertEquals(2, merged1.getLastRow());

            // B: R3-R5
            CellRangeAddress merged2 = sheet.getMergedRegion(1);
            assertEquals(3, merged2.getFirstRow());
            assertEquals(5, merged2.getLastRow());

            // C: R6-R9
            CellRangeAddress merged3 = sheet.getMergedRegion(2);
            assertEquals(6, merged3.getFirstRow());
            assertEquals(9, merged3.getLastRow());
        }

        System.out.println("✓ 多组合并测试通过");
    }

    @Test
    void testSmartMerge_WithMinSpan() throws Exception {
        // 创建测试模板
        byte[] template = createSimpleTemplate();

        // 配置：minSpan=3，至少 3 个相同值才合并
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "groups",
                  "header": {"match": "分组"},
                  "mode": "FILL_DOWN",
                  "merge": {
                    "enabled": true,
                    "minSpan": 3
                  }
                }
              ]
            }
            """;

        // 数据：A-A-B-B-B-C-C
        Map<String, Object> data = new HashMap<>();
        data.put("groups", Arrays.asList("A", "A", "B", "B", "B", "C", "C"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 只有 B 组被合并（3 个）
            assertEquals(1, sheet.getNumMergedRegions(), "应该只有 1 个合并区域");

            CellRangeAddress merged = sheet.getMergedRegion(0);
            assertEquals(3, merged.getFirstRow());
            assertEquals(5, merged.getLastRow());
        }

        System.out.println("✓ 最小合并数测试通过");
    }

    @Test
    void testSmartMerge_MultiColumn() throws Exception {
        // 创建测试模板（2 列）
        byte[] template = createMultiColumnTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "departments",
                  "header": {"match": "部门"},
                  "mode": "FILL_DOWN",
                  "merge": {"enabled": true}
                },
                {
                  "key": "teams",
                  "header": {"match": "团队"},
                  "mode": "FILL_DOWN",
                  "merge": {"enabled": true}
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        // 部门：3 个技术部 + 2 个销售部
        data.put("departments", Arrays.asList("技术部", "技术部", "技术部", "销售部", "销售部"));
        // 团队：2 个 A 组 + 3 个 B 组
        data.put("teams", Arrays.asList("A 组", "A 组", "B 组", "B 组", "B 组"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 4 个合并区域（2 列各 2 个）
            assertEquals(4, sheet.getNumMergedRegions());

            // A 列：技术部（R1-R3），销售部（R4-R5）
            // B 列：A 组（R1-R2），B 组（R3-R5）
        }

        System.out.println("✓ 多列智能合并测试通过");
    }

    @Test
    void testSmartMerge_MixedWithNormalColumn() throws Exception {
        // 创建测试模板
        byte[] template = createMultiColumnTemplate();

        // 部门启用合并，姓名不启用
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "departments",
                  "header": {"match": "部门"},
                  "mode": "FILL_DOWN",
                  "merge": {"enabled": true}
                },
                {
                  "key": "names",
                  "header": {"match": "姓名"},
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        data.put("departments", Arrays.asList("技术部", "技术部", "销售部", "销售部"));
        data.put("names", Arrays.asList("张三", "李四", "王五", "赵六"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 只有 2 个合并区域（部门列）
            assertEquals(2, sheet.getNumMergedRegions());

            // 姓名列应该有 4 个不同的值
            assertEquals("张三", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(1).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(1).getStringCellValue());
            assertEquals("赵六", sheet.getRow(4).getCell(1).getStringCellValue());
        }

        System.out.println("✓ 混合模式测试通过");
    }

    @Test
    void testSmartMerge_NumericValues() throws Exception {
        // 创建测试模板
        byte[] template = createSimpleTemplate();

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "scores",
                  "header": {"match": "分数"},
                  "mode": "FILL_DOWN",
                  "merge": {"enabled": true}
                }
              ]
            }
            """;

        // 数值数据
        Map<String, Object> data = new HashMap<>();
        data.put("scores", Arrays.asList(100, 100, 100, 80, 80, 90));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 2 个合并区域 (100 的 3 个，80 的 2 个)
            assertEquals(2, sheet.getNumMergedRegions(), "应该有 2 个合并区域");

            // 验证数值 - 检查合并区域的第一个单元格
            Cell cell1 = sheet.getRow(1).getCell(3);  // 分数列是第 4 列（索引 3）
            Cell cell4 = sheet.getRow(4).getCell(3);
            Cell cell6 = sheet.getRow(6).getCell(3);

            System.out.println("R1C3 type: " + cell1.getCellType() + ", value: " + cell1);
            System.out.println("R4C3 type: " + cell4.getCellType() + ", value: " + cell4);
            System.out.println("R6C3 type: " + cell6.getCellType() + ", value: " + cell6);

            assertEquals(100.0, cell1.getNumericCellValue(), 0.01);
            assertEquals(80.0, cell4.getNumericCellValue(), 0.01);
            assertEquals(90.0, cell6.getNumericCellValue(), 0.01);
        }

        System.out.println("✓ 数值类型合并测试通过");
    }

    @Test
    void testFixedMerge_RowSpan() throws Exception {
        // 创建测试模板
        byte[] template = createSimpleTemplate();

        // 固定区域合并：每个数据合并 2 行
        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "titles",
                  "header": {"match": "标题"},
                  "mode": "FILL_DOWN",
                  "merge": {
                    "rowSpan": 2,
                    "colSpan": 1
                  }
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        data.put("titles", Arrays.asList("标题 1", "标题 2", "标题 3"));

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);

        // 验证
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 3 个合并区域
            assertEquals(3, sheet.getNumMergedRegions());

            // 每个区域合并 2 行
            CellRangeAddress merged1 = sheet.getMergedRegion(0);
            assertEquals(1, merged1.getFirstRow());
            assertEquals(2, merged1.getLastRow());
        }

        System.out.println("✓ 固定区域合并测试通过");
    }

    // ===== 辅助方法 =====

    private byte[] createSimpleTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("部门");
            headerRow.createCell(1).setCellValue("姓名");
            headerRow.createCell(2).setCellValue("分组");
            headerRow.createCell(3).setCellValue("分数");
            headerRow.createCell(4).setCellValue("标题");

            // 预留数据行
            for (int i = 1; i <= 10; i++) {
                var row = sheet.createRow(i);
                row.createCell(0);
                row.createCell(1);
                row.createCell(2);
                row.createCell(3);
                row.createCell(4);
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createMultiColumnTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");

            // 表头
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("部门");
            headerRow.createCell(1).setCellValue("姓名");
            headerRow.createCell(2).setCellValue("团队");

            // 预留数据行
            for (int i = 1; i <= 10; i++) {
                var row = sheet.createRow(i);
                row.createCell(0);
                row.createCell(1);
                row.createCell(2);
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }
}
