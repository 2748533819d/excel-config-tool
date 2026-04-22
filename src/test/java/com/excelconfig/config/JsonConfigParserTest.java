package com.excelconfig.config;

import com.excelconfig.model.ExcelConfig;
import com.excelconfig.model.ExtractConfig;
import com.excelconfig.model.ExportConfig;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JSON 配置解析器测试
 */
class JsonConfigParserTest {

    private final JsonConfigParser parser = new JsonConfigParser();

    @Test
    void testParseBasicConfig() throws Exception {
        String json = """
            {
              "version": "1.0",
              "templateName": "测试模板",
              "extractions": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "DOWN",
                  "range": { "skipEmpty": true }
                }
              ],
              "exports": [
                {
                  "key": "orderNos",
                  "header": { "match": "订单号" },
                  "mode": "FILL_DOWN"
                }
              ]
            }
            """;

        ExcelConfig config = parser.parse(json);

        assertNotNull(config);
        assertEquals("1.0", config.getVersion());
        assertEquals("测试模板", config.getTemplateName());
        assertEquals(1, config.getExtractions().size());
        assertEquals(1, config.getExports().size());
    }

    @Test
    void testParseExtractConfig() throws Exception {
        String json = """
            {
              "version": "1.0",
              "extractions": [
                {
                  "key": "orderNos",
                  "header": {
                    "match": "订单号",
                    "inRows": [1, 10]
                  },
                  "mode": "DOWN",
                  "range": {
                    "skipEmpty": true,
                    "maxRows": 1000
                  },
                  "parser": {
                    "type": "string"
                  }
                },
                {
                  "key": "amounts",
                  "header": { "match": "金额" },
                  "mode": "DOWN",
                  "parser": {
                    "type": "number",
                    "format": "#,##0.00"
                  }
                }
              ]
            }
            """;

        ExcelConfig config = parser.parse(json);

        assertNotNull(config);
        assertEquals(2, config.getExtractions().size());

        ExtractConfig extract1 = config.getExtractions().get(0);
        assertEquals("orderNos", extract1.getKey());
        assertEquals("订单号", extract1.getHeader().getMatch());
        assertArrayEquals(new int[]{1, 10}, extract1.getHeader().getInRows());
        assertEquals("DOWN", extract1.getMode());
        assertTrue(extract1.getRange().getSkipEmpty());
        assertEquals(1000, extract1.getRange().getMaxRows());
        assertEquals("string", extract1.getParser().getType());
    }

    @Test
    void testParseExportConfig() throws Exception {
        String json = """
            {
              "version": "1.0",
              "exports": [
                {
                  "key": "orders",
                  "header": { "match": "订单号" },
                  "mode": "FILL_TABLE",
                  "columns": [
                    {
                      "key": "orderNo",
                      "header": "订单号",
                      "width": 15
                    },
                    {
                      "key": "amount",
                      "header": "金额",
                      "width": 12,
                      "format": "#,##0.00"
                    }
                  ],
                  "headerStyle": {
                    "bold": true,
                    "background": "#4472C4",
                    "horizontalAlign": "CENTER"
                  },
                  "alternateRows": true,
                  "autoWidth": true
                }
              ]
            }
            """;

        ExcelConfig config = parser.parse(json);

        assertNotNull(config);
        assertEquals(1, config.getExports().size());

        ExportConfig export = config.getExports().get(0);
        assertEquals("orders", export.getKey());
        assertEquals("FILL_TABLE", export.getMode());
        assertEquals(2, export.getColumns().size());
        assertTrue(export.getAlternateRows());
        assertTrue(export.getAutoWidth());
    }

    @Test
    void testParseWithPosition() throws Exception {
        String json = """
            {
              "version": "1.0",
              "extractions": [
                {
                  "key": "title",
                  "position": { "cellRef": "A1" },
                  "mode": "SINGLE"
                }
              ]
            }
            """;

        ExcelConfig excelConfig = parser.parse(json);
        assertNotNull(excelConfig);
        ExtractConfig config = excelConfig.getExtractions().get(0);
        assertEquals("title", config.getKey());
        assertEquals("A1", config.getPosition().getCellRef());
        assertEquals("SINGLE", config.getMode());
    }
}
