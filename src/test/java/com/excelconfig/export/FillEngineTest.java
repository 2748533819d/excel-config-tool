package com.excelconfig.export;

import com.excelconfig.model.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 导出引擎测试
 */
class FillEngineTest {

    private final FillEngine engine = new FillEngine();

    @Test
    void testFill_FillCell() throws IOException {
        // 创建模板
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");
        sheet.createRow(0).createCell(0).setCellValue("标题");

        // 配置
        ExcelConfig config = new ExcelConfig();
        ExportConfig exportConfig = new ExportConfig();
        exportConfig.setKey("title");

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("标题");
        exportConfig.setHeader(headerConfig);

        exportConfig.setMode("FILL_CELL");

        config.getExports().add(exportConfig);

        // 数据
        Map<String, Object> data = new HashMap<>();
        data.put("title", "2024 年销售报表");

        // 执行填充 - 先创建一个临时输入流
        java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
        workbook.write(baos);
        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
        byte[] result = engine.fill(bais, data, config);

        // 验证结果
        assertNotNull(result);
        assertTrue(result.length > 0);

        // 读取结果验证内容
        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet resultSheet = resultWorkbook.getSheetAt(0);
            Row row = resultSheet.getRow(1);
            assertNotNull(row);
            Cell cell = row.getCell(0);
            assertNotNull(cell);
            assertEquals("2024 年销售报表", cell.getStringCellValue());
        }
    }

    @Test
    void testFill_FillDown() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        // 表头
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("订单号");

        // 配置
        ExcelConfig config = new ExcelConfig();
        ExportConfig exportConfig = new ExportConfig();
        exportConfig.setKey("orderNos");

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("订单号");
        exportConfig.setHeader(headerConfig);

        exportConfig.setMode("FILL_DOWN");
        config.getExports().add(exportConfig);

        // 数据
        Map<String, Object> data = new HashMap<>();
        data.put("orderNos", Arrays.asList("ORD001", "ORD002", "ORD003", "ORD004"));

        // 执行填充
        java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
        workbook.write(baos);
        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
        byte[] result = engine.fill(bais, data, config);

        // 验证结果
        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet resultSheet = resultWorkbook.getSheetAt(0);

            assertEquals("ORD001", resultSheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("ORD002", resultSheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("ORD003", resultSheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals("ORD004", resultSheet.getRow(4).getCell(0).getStringCellValue());
        }
    }

    @Test
    void testFill_FillTable() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        // 创建表头行（用于匹配定位）
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("订单号");

        // 配置
        ExcelConfig config = new ExcelConfig();
        ExportConfig exportConfig = new ExportConfig();
        exportConfig.setKey("orders");

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("订单号");
        exportConfig.setHeader(headerConfig);

        exportConfig.setMode("FILL_TABLE");

        // 列配置
        ColumnConfig col1 = new ColumnConfig();
        col1.setKey("orderNo");
        col1.setHeader("订单号");
        col1.setWidth(15);

        ColumnConfig col2 = new ColumnConfig();
        col2.setKey("amount");
        col2.setHeader("金额");
        col2.setWidth(12);

        exportConfig.setColumns(Arrays.asList(col1, col2));
        config.getExports().add(exportConfig);

        // 数据
        Map<String, Object> data = new HashMap<>();
        List<Map<String, Object>> orders = Arrays.asList(
            Map.of("orderNo", "ORD001", "amount", 100.00),
            Map.of("orderNo", "ORD002", "amount", 200.00)
        );
        data.put("orders", orders);

        // 执行填充
        java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
        workbook.write(baos);
        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
        byte[] result = engine.fill(bais, data, config);

        // 验证结果
        try (Workbook resultWorkbook = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet resultSheet = resultWorkbook.getSheetAt(0);

            // 验证表头
            Row resultHeaderRow = resultSheet.getRow(0);
            assertNotNull(resultHeaderRow, "表头行应该存在");
            Cell headerCell0 = resultHeaderRow.getCell(0);
            assertNotNull(headerCell0, "表头第 1 列应该存在");
            assertEquals("订单号", headerCell0.getStringCellValue());

            Cell headerCell1 = resultHeaderRow.getCell(1);
            assertNotNull(headerCell1, "表头第 2 列应该存在");
            assertEquals("金额", headerCell1.getStringCellValue());

            // 验证数据
            Row dataRow1 = resultSheet.getRow(1);
            assertNotNull(dataRow1, "数据第 1 行应该存在");
            Cell dataCell1_0 = dataRow1.getCell(0);
            assertNotNull(dataCell1_0, "数据第 1 行第 1 列应该存在");
            assertEquals("ORD001", dataCell1_0.getStringCellValue());

            Cell dataCell1_1 = dataRow1.getCell(1);
            assertNotNull(dataCell1_1, "数据第 1 行第 2 列应该存在");
            assertEquals(100.00, dataCell1_1.getNumericCellValue(), 0.01);

            Row dataRow2 = resultSheet.getRow(2);
            assertNotNull(dataRow2, "数据第 2 行应该存在");
            assertEquals("ORD002", dataRow2.getCell(0).getStringCellValue());
            assertEquals(200.00, dataRow2.getCell(1).getNumericCellValue(), 0.01);
        }
    }
}
