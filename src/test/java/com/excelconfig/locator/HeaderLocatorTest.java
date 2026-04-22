package com.excelconfig.locator;

import com.excelconfig.model.HeaderConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 表头定位器测试
 */
class HeaderLocatorTest {

    private final HeaderLocator locator = new HeaderLocator();

    @Test
    void testLocate_HeaderInFirstRow() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        // 创建表头：A1="订单号", B1="金额"
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("订单号");
        headerRow.createCell(1).setCellValue("金额");

        HeaderConfig config = new HeaderConfig();
        config.setMatch("订单号");

        HeaderPosition pos = locator.locate(sheet, config);

        assertEquals(0, pos.getRow());
        assertEquals(0, pos.getColumn());
        assertEquals("A1", pos.toString());

        workbook.close();
    }

    @Test
    void testLocate_HeaderInRange() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        // 第 0 行：标题
        sheet.createRow(0).createCell(0).setCellValue("公司报表");
        // 第 1 行：副标题
        sheet.createRow(1).createCell(0).setCellValue("2024 年 1 月");
        // 第 2 行：表头
        Row headerRow = sheet.createRow(2);
        headerRow.createCell(0).setCellValue("订单号");
        headerRow.createCell(1).setCellValue("金额");

        HeaderConfig config = new HeaderConfig();
        config.setMatch("订单号");
        config.setInRows(new int[]{1, 5});  // 在第 1-5 行搜索

        HeaderPosition pos = locator.locate(sheet, config);

        assertEquals(2, pos.getRow());
        assertEquals(0, pos.getColumn());

        workbook.close();
    }

    @Test
    void testLocate_HeaderNotFound() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("订单号");

        HeaderConfig config = new HeaderConfig();
        config.setMatch("不存在的表头");

        assertThrows(HeaderNotFoundException.class, () -> {
            locator.locate(sheet, config);
        });

        workbook.close();
    }

    @Test
    void testLocate_NumericHeader() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("2024");  // 数字类型

        HeaderConfig config = new HeaderConfig();
        config.setMatch("2024");

        HeaderPosition pos = locator.locate(sheet, config);

        assertEquals(0, pos.getRow());
        assertEquals(0, pos.getColumn());

        workbook.close();
    }
}
