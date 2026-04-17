package com.excelconfig.extract;

import com.excelconfig.model.ExcelConfig;
import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.HeaderConfig;
import com.excelconfig.model.RangeConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 提取引擎测试
 */
class ExtractEngineTest {

    private final ExtractEngine engine = new ExtractEngine();

    @Test
    void testExtract_DownMode() throws IOException {
        // 创建测试 Excel
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        // 表头：A1="订单号"
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("订单号");

        // 数据：A2-A5
        sheet.createRow(1).createCell(0).setCellValue("ORD001");
        sheet.createRow(2).createCell(0).setCellValue("ORD002");
        sheet.createRow(3).createCell(0).setCellValue("ORD003");
        sheet.createRow(4).createCell(0).setCellValue("ORD004");

        // 创建配置
        ExcelConfig config = new ExcelConfig();
        ExtractConfig extractConfig = new ExtractConfig();
        extractConfig.setKey("orderNos");

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("订单号");
        extractConfig.setHeader(headerConfig);

        extractConfig.setMode("DOWN");

        RangeConfig rangeConfig = new RangeConfig();
        rangeConfig.setSkipEmpty(true);
        extractConfig.setRange(rangeConfig);

        config.getExtractions().add(extractConfig);

        // 执行提取
        List<Object> orderNos = engine.extract(sheet, extractConfig);

        // 验证结果
        assertNotNull(orderNos);
        assertEquals(4, orderNos.size());
        assertEquals("ORD001", orderNos.get(0));
        assertEquals("ORD002", orderNos.get(1));
        assertEquals("ORD003", orderNos.get(2));
        assertEquals("ORD004", orderNos.get(3));

        workbook.close();
    }

    @Test
    void testExtract_SingleMode() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        sheet.createRow(0).createCell(0).setCellValue("标题");
        sheet.createRow(1).createCell(0).setCellValue("2024 年报表");

        ExcelConfig config = new ExcelConfig();
        ExtractConfig extractConfig = new ExtractConfig();
        extractConfig.setKey("title");

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("标题");
        extractConfig.setHeader(headerConfig);

        extractConfig.setMode("SINGLE");
        config.getExtractions().add(extractConfig);

        List<Object> titles = engine.extract(sheet, extractConfig);

        assertNotNull(titles);
        assertEquals(1, titles.size());
        assertEquals("2024 年报表", titles.get(0));

        workbook.close();
    }

    @Test
    void testExtract_DownModeWithEmptyRow() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        // 表头
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("订单号");

        // 数据（中间有空行）
        sheet.createRow(1).createCell(0).setCellValue("ORD001");
        sheet.createRow(2).createCell(0).setCellValue("ORD002");
        sheet.createRow(3);  // 空行
        sheet.createRow(4).createCell(0).setCellValue("ORD003");

        ExcelConfig config = new ExcelConfig();
        ExtractConfig extractConfig = new ExtractConfig();
        extractConfig.setKey("orderNos");

        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("订单号");
        extractConfig.setHeader(headerConfig);

        extractConfig.setMode("DOWN");

        RangeConfig rangeConfig = new RangeConfig();
        rangeConfig.setSkipEmpty(true);
        extractConfig.setRange(rangeConfig);

        config.getExtractions().add(extractConfig);

        List<Object> orderNos = engine.extract(sheet, extractConfig);

        // 应该跳过空行，继续读取
        assertEquals(3, orderNos.size());
        assertEquals("ORD001", orderNos.get(0));
        assertEquals("ORD002", orderNos.get(1));
        assertEquals("ORD003", orderNos.get(2));

        workbook.close();
    }
}
