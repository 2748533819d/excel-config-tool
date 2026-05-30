package com.excelconfig.util;

import com.excelconfig.model.StyleConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * CellStyle 缓存 - 避免每单元格重复创建样式对象
 *
 * <p>POI 的 CellStyle 在 Workbook 级别共享，创建开销较大。
 * 对于大数据量填充，应该复用相同配置的 CellStyle 实例。
 * 本缓存以 StyleConfig 的哈希值为 key，确保相同配置返回相同实例。</p>
 */
public class StyleCache {

    private static final Logger log = LoggerFactory.getLogger(StyleCache.class);

    private final Map<StyleKey, CellStyle> cache = new HashMap<>();
    private final Workbook workbook;

    public StyleCache(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 获取或创建缓存的 CellStyle
     *
     * @param style 样式配置（允许 null，返回 null）
     * @return 缓存的 CellStyle 实例，style 为 null 时返回 null
     */
    public CellStyle getOrCreateStyle(StyleConfig style) {
        if (style == null) {
            return null;
        }

        StyleKey key = new StyleKey(style);
        return cache.computeIfAbsent(key, k -> createStyle(style));
    }

    /**
     * 根据配置创建 CellStyle
     */
    private CellStyle createStyle(StyleConfig style) {
        CellStyle cellStyle = workbook.createCellStyle();

        // 水平对齐
        if (style.getHorizontalAlign() != null) {
            switch (style.getHorizontalAlign().toUpperCase()) {
                case "LEFT":
                    cellStyle.setAlignment(HorizontalAlignment.LEFT);
                    break;
                case "CENTER":
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case "RIGHT":
                    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                default:
                    break;
            }
        }

        // 垂直对齐
        if (style.getVerticalAlign() != null) {
            switch (style.getVerticalAlign().toUpperCase()) {
                case "TOP":
                    cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
                    break;
                case "CENTER":
                    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    break;
                case "BOTTOM":
                    cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    break;
                default:
                    break;
            }
        }

        // 数字格式
        if (style.getFormat() != null) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(style.getFormat()));
        }

        // 背景颜色
        if (style.getBackground() != null) {
            applyBackgroundColor(cellStyle, style.getBackground());
        }

        // 字体（合并 bold + fontColor 到同一个 Font 对象，避免后者覆盖前者）
        boolean hasBold = Boolean.TRUE.equals(style.getBold());
        boolean hasFontColor = style.getFontColor() != null;
        if (hasBold || hasFontColor) {
            Font font = workbook.createFont();
            if (hasBold) {
                font.setBold(true);
            }
            if (hasFontColor && workbook instanceof XSSFWorkbook) {
                applyFontColorToFont(font, style.getFontColor());
            }
            cellStyle.setFont(font);
        }

        return cellStyle;
    }

    private void applyBackgroundColor(CellStyle cellStyle, String colorHex) {
        if (workbook instanceof XSSFWorkbook) {
            try {
                String hex = colorHex.startsWith("#") ? colorHex.substring(1) : colorHex;
                if (hex.length() == 6) {
                    int r = Integer.parseInt(hex.substring(0, 2), 16);
                    int g = Integer.parseInt(hex.substring(2, 4), 16);
                    int b = Integer.parseInt(hex.substring(4, 6), 16);
                    XSSFColor xssfColor = new XSSFColor(new java.awt.Color(r, g, b), new DefaultIndexedColorMap());
                    cellStyle.setFillForegroundColor(xssfColor);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    return;
                }
            } catch (Exception e) {
                log.warn("解析背景颜色失败 [{}]，使用默认灰色: {}", colorHex, e.getMessage());
            }
        }
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void applyFontColorToFont(Font font, String colorHex) {
        if (font instanceof XSSFFont xssfFont) {
            try {
                String hex = colorHex.startsWith("#") ? colorHex.substring(1) : colorHex;
                if (hex.length() == 6) {
                    int r = Integer.parseInt(hex.substring(0, 2), 16);
                    int g = Integer.parseInt(hex.substring(2, 4), 16);
                    int b = Integer.parseInt(hex.substring(4, 6), 16);
                    xssfFont.setColor(new XSSFColor(new java.awt.Color(r, g, b), new DefaultIndexedColorMap()));
                }
            } catch (Exception e) {
                log.warn("解析字体颜色失败 [{}]: {}", colorHex, e.getMessage());
            }
        }
    }

    /**
     * 样式键 - 基于 StyleConfig 属性计算 hash，用于缓存查找
     */
    private static class StyleKey {
        private final boolean bold;
        private final String background;
        private final String fontColor;
        private final Integer fontSize;
        private final String format;
        private final String horizontalAlign;
        private final String verticalAlign;

        StyleKey(StyleConfig config) {
            this.bold = Boolean.TRUE.equals(config.getBold());
            this.background = config.getBackground();
            this.fontColor = config.getFontColor();
            this.fontSize = config.getFontSize();
            this.format = config.getFormat();
            this.horizontalAlign = config.getHorizontalAlign();
            this.verticalAlign = config.getVerticalAlign();
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof StyleKey)) return false;
            StyleKey styleKey = (StyleKey) o;
            return bold == styleKey.bold
                && Objects.equals(background, styleKey.background)
                && Objects.equals(fontColor, styleKey.fontColor)
                && Objects.equals(fontSize, styleKey.fontSize)
                && Objects.equals(format, styleKey.format)
                && Objects.equals(horizontalAlign, styleKey.horizontalAlign)
                && Objects.equals(verticalAlign, styleKey.verticalAlign);
        }

        @Override
        public int hashCode() {
            return Objects.hash(bold, background, fontColor, fontSize, format, horizontalAlign, verticalAlign);
        }
    }
}
