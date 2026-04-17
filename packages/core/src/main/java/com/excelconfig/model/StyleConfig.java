package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 样式配置
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class StyleConfig {

    /**
     * 是否加粗
     */
    private Boolean bold;

    /**
     * 背景颜色（十六进制）
     */
    private String background;

    /**
     * 字体颜色（十六进制）
     */
    private String fontColor;

    /**
     * 字体大小
     */
    private Integer fontSize;

    /**
     * 数字格式
     */
    private String format;

    /**
     * 水平对齐：LEFT, CENTER, RIGHT
     */
    private String horizontalAlign;

    /**
     * 垂直对齐：TOP, CENTER, BOTTOM
     */
    private String verticalAlign;

    public StyleConfig() {
    }

    public Boolean getBold() {
        return bold;
    }

    public void setBold(Boolean bold) {
        this.bold = bold;
    }

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public String getFontColor() {
        return fontColor;
    }

    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    public Integer getFontSize() {
        return fontSize;
    }

    public void setFontSize(Integer fontSize) {
        this.fontSize = fontSize;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public String getHorizontalAlign() {
        return horizontalAlign;
    }

    public void setHorizontalAlign(String horizontalAlign) {
        this.horizontalAlign = horizontalAlign;
    }

    public String getVerticalAlign() {
        return verticalAlign;
    }

    public void setVerticalAlign(String verticalAlign) {
        this.verticalAlign = verticalAlign;
    }
}
