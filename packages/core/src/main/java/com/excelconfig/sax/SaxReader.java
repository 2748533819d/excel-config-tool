package com.excelconfig.sax;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * SAX 流式读取器
 *
 * 用于大文件的内存优化读取，基于 Apache POI 的事件模型（Event API）
 */
public class SaxReader {

    /**
     * 从 Excel 文件中流式读取指定 Sheet 的数据
     */
    public void read(InputStream inputStream, int sheetIndex, RowHandler rowHandler) throws Exception {
        // 打开 OPC 包
        org.apache.poi.openxml4j.opc.OPCPackage pkg = org.apache.poi.openxml4j.opc.OPCPackage.open(inputStream);
        XSSFReader xssfReader = new XSSFReader(pkg);

        // 获取共享字符串表（只读版本）
        ReadOnlySharedStringsTable sharedStrings = new ReadOnlySharedStringsTable(pkg);
        StylesTable styles = xssfReader.getStylesTable();

        // 获取指定 Sheet 的输入流
        Iterator<InputStream> sheets = xssfReader.getSheetsData();
        InputStream sheetInputStream = null;

        try {
            // 定位到指定 Sheet
            int currentIndex = 0;
            while (sheets.hasNext()) {
                sheetInputStream = sheets.next();
                if (currentIndex == sheetIndex) {
                    break;
                }
                sheetInputStream.close();
                currentIndex++;
            }

            if (sheetInputStream == null || currentIndex != sheetIndex) {
                throw new IllegalArgumentException("Sheet index " + sheetIndex + " not found");
            }

            // 创建 XML 读取器
            javax.xml.parsers.SAXParserFactory factory = javax.xml.parsers.SAXParserFactory.newInstance();
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            javax.xml.parsers.SAXParser saxParser = factory.newSAXParser();
            org.xml.sax.XMLReader sheetParser = saxParser.getXMLReader();

            // 创建自定义 Handler
            SimpleSheetHandler handler = new SimpleSheetHandler(sharedStrings, styles, rowHandler);

            sheetParser.setContentHandler(handler);
            sheetParser.parse(new org.xml.sax.InputSource(sheetInputStream));

        } finally {
            if (sheetInputStream != null) {
                sheetInputStream.close();
            }
            pkg.close();
        }
    }

    /**
     * 流式读取所有数据到内存
     */
    public List<List<String>> readAll(InputStream inputStream, int sheetIndex) throws Exception {
        List<List<String>> result = new ArrayList<>();

        read(inputStream, sheetIndex, (rowNum, cells) -> result.add(cells));

        return result;
    }
}