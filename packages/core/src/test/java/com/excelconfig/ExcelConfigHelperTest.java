package com.excelconfig;

import com.excelconfig.model.ExcelConfig;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * ExcelConfigHelper 门面类测试
 */
class ExcelConfigHelperTest {

    @Test
    void testReadChain() {
        // 测试链式调用
        ExcelConfigHelper helper = ExcelConfigHelper.read("test.xlsx");
        assertNotNull(helper);
    }

    @Test
    void testWriteChain() {
        ExcelConfigHelper helper = ExcelConfigHelper.write("test.xlsx");
        assertNotNull(helper);
    }

    @Test
    void testConfigMethods() throws Exception {
        String configJson = "{" +
            "\"extractions\": [" +
            "  {\"key\": \"name\", \"header\": {\"match\": \"姓名\"}, \"mode\": \"DOWN\"}" +
            "]" +
            "}";

        ExcelConfigHelper helper = ExcelConfigHelper.read("test.xlsx");
        assertNotNull(helper.configJson(configJson));
    }

    @Test
    void testDataMethod() {
        Map<String, Object> data = new HashMap<>();
        ExcelConfigHelper helper = ExcelConfigHelper.write("test.xlsx");
        assertNotNull(helper.data(data));
    }

    @Test
    void testStaticReadMethods() {
        // 测试静态 read 方法
        assertNotNull(ExcelConfigHelper.read("test.xlsx"));
        assertNotNull(ExcelConfigHelper.read(new File("test.xlsx")));
        assertNotNull(ExcelConfigHelper.read(new ByteArrayInputStream(new byte[0])));
    }

    @Test
    void testStaticWriteMethods() {
        // 测试静态 write 方法
        assertNotNull(ExcelConfigHelper.write("test.xlsx"));
        assertNotNull(ExcelConfigHelper.write(new File("test.xlsx")));
        assertNotNull(ExcelConfigHelper.write(new ByteArrayInputStream(new byte[0])));
    }
}
