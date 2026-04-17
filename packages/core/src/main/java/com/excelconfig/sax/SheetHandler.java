package com.excelconfig.sax;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * 简单的 Sheet 处理器 - 处理 SAX 事件并解析单元格数据
 */
class SimpleSheetHandler extends DefaultHandler {

    private final ReadOnlySharedStringsTable sharedStrings;
    private final StylesTable styles;
    private final RowHandler rowHandler;

    private int currentRowNum = -1;
    private List<String> currentRow;
    private String lastContents = "";
    private boolean cellIsOpen = false;
    private int currentColumn = -1;
    private String cellType = null;

    public SimpleSheetHandler(ReadOnlySharedStringsTable sharedStrings, StylesTable styles, RowHandler rowHandler) {
        this.sharedStrings = sharedStrings;
        this.styles = styles;
        this.rowHandler = rowHandler;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        // 行开始
        if ("row".equals(name)) {
            String rowNumStr = attributes.getValue("r");
            if (rowNumStr != null) {
                currentRowNum = Integer.parseInt(rowNumStr) - 1;
            } else {
                currentRowNum++;
            }
            currentRow = new ArrayList<>();
            currentColumn = -1;
        }

        // 单元格开始
        if ("c".equals(name)) {
            cellIsOpen = true;
            lastContents = "";
            currentColumn++;
            cellType = attributes.getValue("t");

            // 获取单元格引用（如 A1, B1 等）
            String cellRef = attributes.getValue("r");
            if (cellRef != null) {
                currentColumn = parseColumnIndex(cellRef);
            }
        }

        // 单元格值开始
        if ("v".equals(name)) {
            lastContents = "";
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        // 单元格结束
        if ("c".equals(name)) {
            cellIsOpen = false;

            // 填充空单元格
            while (currentRow.size() <= currentColumn) {
                currentRow.add("");
            }

            // 设置单元格值
            currentRow.set(currentColumn, getCellValue(lastContents, cellType));
        }

        // 行结束
        if ("row".equals(name)) {
            if (rowHandler != null && currentRow != null) {
                rowHandler.handleRow(currentRowNum, new ArrayList<>(currentRow));
            }
            currentRow = null;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        if (cellIsOpen) {
            lastContents += new String(ch, start, length);
        }
    }

    private String getCellValue(String value, String type) {
        if (value == null || value.isEmpty()) {
            return "";
        }

        // 共享字符串类型
        if ("s".equals(type)) {
            try {
                int index = Integer.parseInt(value);
                org.apache.poi.ss.usermodel.RichTextString sharedString = sharedStrings.getItemAt(index);
                return sharedString.getString();
            } catch (Exception e) {
                return value;
            }
        }

        // 内联字符串类型
        if ("inlineStr".equals(type)) {
            return value;
        }

        // 布尔类型
        if ("b".equals(type)) {
            return "1".equals(value) ? "TRUE" : "FALSE";
        }

        // 其他类型（数字、日期等）直接返回
        return value;
    }

    /**
     * 从单元格引用中解析列索引（如 A1 -> 0, B2 -> 1）
     */
    private int parseColumnIndex(String cellReference) {
        if (cellReference == null || cellReference.isEmpty()) {
            return 0;
        }

        int columnIndex = 0;
        int i = 0;

        // 解析字母部分
        while (i < cellReference.length() && Character.isUpperCase(cellReference.charAt(i))) {
            columnIndex = columnIndex * 26 + (cellReference.charAt(i) - 'A' + 1);
            i++;
        }

        return columnIndex - 1;
    }
}
