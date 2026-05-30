package com.excelconfig;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
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
        byte[] template = createSimpleTemplate("部门");

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
        byte[] template = createSimpleTemplate("部门");

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
        byte[] template = createSimpleTemplate("分组");

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
        byte[] template = createSimpleTemplate("分组");

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
        byte[] template = createSimpleTemplate("分数");

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

            // 验证数值 - 检查合并区域的第一个单元格（单列表格，列索引 0）
            Cell cell1 = sheet.getRow(1).getCell(0);
            Cell cell4 = sheet.getRow(4).getCell(0);
            Cell cell6 = sheet.getRow(6).getCell(0);

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
        byte[] template = createSimpleTemplate("标题");

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

    // ========== FILL_TABLE 合并测试 ==========

    @Test
    void testFillTable_SmartMerge_ColumnLevel() throws Exception {
        // 方案A：FILL_TABLE 列级智能合并
        // 部门列有相同值，按部门合并；姓名列不合并
        byte[] template = createTableTemplate("部门", "姓名");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "data",
                  "header": {"match": "部门"},
                  "mode": "FILL_TABLE",
                  "columns": [
                    {
                      "key": "dept",
                      "header": "部门",
                      "merge": { "enabled": true }
                    },
                    {
                      "key": "name",
                      "header": "姓名"
                    }
                  ]
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = Arrays.asList(
            Map.of("dept", "技术部", "name", "张三"),
            Map.of("dept", "技术部", "name", "李四"),
            Map.of("dept", "销售部", "name", "王五"),
            Map.of("dept", "销售部", "name", "赵六")
        );
        data.put("data", rows);

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);
        saveOutput(result, "testFillTable_SmartMerge_ColumnLevel");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 部门列：2 个合并区域
            assertEquals(2, sheet.getNumMergedRegions(), "应该有 2 个合并区域");

            // 第一个合并区域：R1-R2 合并（技术部）
            CellRangeAddress m1 = sheet.getMergedRegion(0);
            assertEquals(1, m1.getFirstRow());
            assertEquals(2, m1.getLastRow());
            assertEquals(0, m1.getFirstColumn());
            assertEquals(0, m1.getLastColumn());

            // 第二个合并区域：R3-R4 合并（销售部）
            CellRangeAddress m2 = sheet.getMergedRegion(1);
            assertEquals(3, m2.getFirstRow());
            assertEquals(4, m2.getLastRow());
            assertEquals(0, m2.getFirstColumn());

            // 姓名列不合并，4 行都有值
            assertEquals("张三", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(1).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(1).getStringCellValue());
            assertEquals("赵六", sheet.getRow(4).getCell(1).getStringCellValue());
        }

        System.out.println("✓ FILL_TABLE 列级智能合并测试通过");
    }

    @Test
    void testFillTable_SmartMerge_MinSpan() throws Exception {
        // 方案A：minSpan 过滤，不满足最少合并数的列不合并
        byte[] template = createTableTemplate("部门");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "data",
                  "header": {"match": "部门"},
                  "mode": "FILL_TABLE",
                  "columns": [
                    {
                      "key": "dept",
                      "header": "部门",
                      "merge": { "enabled": true, "minSpan": 3 }
                    }
                  ]
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = Arrays.asList(
            Map.of("dept", "技术部"),
            Map.of("dept", "技术部"),
            Map.of("dept", "销售部")
        );
        data.put("data", rows);

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);
        saveOutput(result, "testFillTable_SmartMerge_MinSpan");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);
            // minSpan=3，只有 2 个技术部不满足条件，不应该合并
            assertEquals(0, sheet.getNumMergedRegions(), "minSpan=3 时不应合并");
        }

        System.out.println("✓ FILL_TABLE 列级智能合并 minSpan 测试通过");
    }

    @Test
    void testFillTable_ColSpan_HeaderMerge() throws Exception {
        // 方案C：表头跨列合并
        byte[] template = createTableTemplate("地址", "列2", "列3");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "data",
                  "header": {"match": "地址"},
                  "mode": "FILL_TABLE",
                  "columns": [
                    {
                      "key": "address",
                      "header": "地址信息",
                      "merge": { "colSpan": 3 }
                    }
                  ]
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = Arrays.asList(
            Map.of("address", "广东省广州市天河区")
        );
        data.put("data", rows);

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);
        saveOutput(result, "testFillTable_ColSpan_HeaderMerge");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 表头合并区域 + 数据行跨列合并区域
            assertTrue(sheet.getNumMergedRegions() >= 1, "至少 1 个合并区域");
            // 第一个区域是表头合并（row 0, cols 0-2）
            CellRangeAddress m1 = sheet.getMergedRegion(0);
            assertEquals(0, m1.getFirstRow(), "表头在 row 0");
            assertEquals(0, m1.getLastRow());
            assertEquals(0, m1.getFirstColumn());
            assertEquals(2, m1.getLastColumn(), "表头跨 3 列");

            // 表头内容
            assertEquals("地址信息", sheet.getRow(0).getCell(0).getStringCellValue());
            // 表头跨列的部分应为空
            assertEquals(CellType.BLANK, sheet.getRow(0).getCell(1).getCellType());
            assertEquals(CellType.BLANK, sheet.getRow(0).getCell(2).getCellType());
        }

        System.out.println("✓ FILL_TABLE 表头跨列合并测试通过");
    }

    @Test
    void testFillTable_ColSpan_DataMerge() throws Exception {
        // 方案C：数据行跨列合并 + Smart Merge 纵向合并
        byte[] template = createTableTemplate("备注", "列2", "列3", "列4");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "data",
                  "header": {"match": "备注"},
                  "mode": "FILL_TABLE",
                  "columns": [
                    {
                      "key": "remark",
                      "header": "备注信息",
                      "merge": { "colSpan": 3, "enabled": true }
                    }
                  ]
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = Arrays.asList(
            Map.of("remark", "待处理"),
            Map.of("remark", "待处理"),
            Map.of("remark", "已完成")
        );
        data.put("data", rows);

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);
        saveOutput(result, "testFillTable_ColSpan_DataMerge");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 1 个表头合并 + 2 个数据合并（智能合并：待处理行合并，已完成单行不合并）
            assertTrue(sheet.getNumMergedRegions() >= 2, "至少 2 个合并区域（表头 + 数据 smart merge）");

            // 验证表头跨列
            assertEquals("备注信息", sheet.getRow(0).getCell(0).getStringCellValue());

            // 验证数据跨列：row1 和 row2 合并（待处理），row3 单独
            // row1 的 3 列都应跨列到第一列有值
            Cell cellR1 = sheet.getRow(1).getCell(0);
            assertEquals("待处理", cellR1.getStringCellValue());
            // 跨列部分为空
            assertEquals(CellType.BLANK, sheet.getRow(1).getCell(1).getCellType());
            assertEquals(CellType.BLANK, sheet.getRow(1).getCell(2).getCellType());
            // 最后一个单元格没跨列
            assertEquals("已完成", sheet.getRow(3).getCell(0).getStringCellValue());
        }

        System.out.println("✓ FILL_TABLE 数据跨列 + 智能合并测试通过");
    }

    @Test
    void testFillTable_MixedColumns_ColSpanAndNormal() throws Exception {
        // 方案A+C混合：跨列合并列 + 普通列共存
        byte[] template = createTableTemplate("地址", "姓名");

        String configJson = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "data",
                  "header": {"match": "地址"},
                  "mode": "FILL_TABLE",
                  "columns": [
                    {
                      "key": "address",
                      "header": "地址",
                      "merge": { "colSpan": 3, "enabled": true }
                    },
                    {
                      "key": "name",
                      "header": "姓名"
                    }
                  ]
                }
              ]
            }
            """;

        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> rows = Arrays.asList(
            Map.of("address", "广州", "name", "张三"),
            Map.of("address", "广州", "name", "李四"),
            Map.of("address", "深圳", "name", "王五")
        );
        data.put("data", rows);

        ExcelConfigService service = new ExcelConfigService();
        byte[] result = service.fill(new ByteArrayInputStream(template), data, configJson);
        saveOutput(result, "testFillTable_MixedColumns_ColSpanAndNormal");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet sheet = workbook.getSheetAt(0);

            // 表头合并(1)：地址跨 3 列
            // 地址列 smart merge(1)：广州 R1-2 合并，深圳单列不合并
            // 所以至少 2 个合并区域
            assertTrue(sheet.getNumMergedRegions() >= 2, "至少 2 个合并区域");

            // 姓名列（物理列位置 = 3）：不受影响
            assertEquals("张三", sheet.getRow(1).getCell(3).getStringCellValue());
            assertEquals("李四", sheet.getRow(2).getCell(3).getStringCellValue());
            assertEquals("王五", sheet.getRow(3).getCell(3).getStringCellValue());
        }

        System.out.println("✓ FILL_TABLE 混合列测试通过");
    }

    // ===== 辅助方法 =====

    /** 测试输出文件保存目录 */
    private static final Path OUTPUT_DIR = Path.of("/Users/huangzhenzhen/Documents/excel-test/未命名文件夹");

    static {
        try {
            Files.createDirectories(OUTPUT_DIR);
        } catch (Exception ignored) {
        }
    }

    private void saveOutput(byte[] data, String name) {
        try {
            Path file = OUTPUT_DIR.resolve(name + ".xlsx");
            Files.write(file, data);
            System.out.println("已保存: " + file);
        } catch (Exception e) {
            System.err.println("保存失败: " + e.getMessage());
        }
    }

    /**
     * 创建适用于 FILL_TABLE 的模板：可指定表头文字
     */
    private byte[] createTableTemplate(String... headers) throws Exception {
        int colCount = headers.length;
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");
            var headerRow = sheet.createRow(0);
            for (int i = 0; i < colCount; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }
            // 预留 20 行数据空间
            for (int i = 1; i <= 20; i++) {
                var row = sheet.createRow(i);
                for (int c = 0; c < colCount; c++) {
                    row.createCell(c);
                }
            }
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    private byte[] createSimpleTemplate(String header) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Test");
            var headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue(header);
            // 预留 20 行数据空间
            for (int i = 1; i <= 20; i++) {
                var row = sheet.createRow(i);
                row.createCell(0);
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
