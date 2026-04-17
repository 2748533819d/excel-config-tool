package com.excelconfig.extract;

import com.excelconfig.model.ParserConfig;
import com.excelconfig.spi.CellParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

/**
 * 默认单元格解析器
 */
public class DefaultCellParser implements CellParser {

    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
    private static final String DEFAULT_NUMBER_FORMAT = "#,##0.##";

    @Override
    public Object parse(Cell cell, ParserConfig config) {
        if (cell == null) {
            return null;
        }

        String type = config != null ? config.getType() : "string";

        switch (type) {
            case "string":
                return parseString(cell);
            case "number":
                return parseNumber(cell, config);
            case "date":
                return parseDate(cell, config);
            case "boolean":
                return parseBoolean(cell);
            default:
                return parseString(cell);
        }
    }

    private String parseString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double num = cell.getNumericCellValue();
                    if (num == (long) num) {
                        return String.valueOf((long) num);
                    }
                    return String.valueOf(num);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            default:
                return null;
        }
    }

    private Number parseNumber(Cell cell, ParserConfig config) {
        switch (cell.getCellType()) {
            case NUMERIC:
                double value = cell.getNumericCellValue();
                if (value == (long) value) {
                    return (long) value;
                }
                return value;
            case STRING:
                String strValue = cell.getStringCellValue();
                try {
                    if (strValue.contains(".")) {
                        return Double.parseDouble(strValue);
                    } else {
                        return Long.parseLong(strValue);
                    }
                } catch (NumberFormatException e) {
                    return null;
                }
            default:
                return null;
        }
    }

    private java.util.Date parseDate(Cell cell, ParserConfig config) {
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return null;
            case STRING:
                String dateStr = cell.getStringCellValue();
                String pattern = config != null && config.getFormat() != null
                    ? config.getFormat()
                    : DEFAULT_DATE_FORMAT;
                try {
                    SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                    return sdf.parse(dateStr);
                } catch (Exception e) {
                    return null;
                }
            default:
                return null;
        }
    }

    private Boolean parseBoolean(Cell cell) {
        switch (cell.getCellType()) {
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                return cell.getNumericCellValue() != 0;
            case STRING:
                String str = cell.getStringCellValue().toLowerCase();
                return "true".equals(str) || "yes".equals(str) || "1".equals(str);
            default:
                return null;
        }
    }
}
