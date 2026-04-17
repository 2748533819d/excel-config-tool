package com.excelconfig.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

/**
 * 配置根类
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class ExcelConfig {

    /**
     * 配置版本
     */
    private String version;

    /**
     * 模板名称
     */
    private String templateName;

    /**
     * 导入配置列表
     */
    @JsonProperty("extractions")
    private java.util.List<ExtractConfig> extractions;

    /**
     * 导出配置列表
     */
    @JsonProperty("exports")
    private java.util.List<ExportConfig> exports;

    public ExcelConfig() {
        this.extractions = new java.util.ArrayList<>();
        this.exports = new java.util.ArrayList<>();
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public String getTemplateName() {
        return templateName;
    }

    public void setTemplateName(String templateName) {
        this.templateName = templateName;
    }

    public java.util.List<ExtractConfig> getExtractions() {
        return extractions;
    }

    public void setExtractions(java.util.List<ExtractConfig> extractions) {
        this.extractions = extractions;
    }

    public java.util.List<ExportConfig> getExports() {
        return exports;
    }

    public void setExports(java.util.List<ExportConfig> exports) {
        this.exports = exports;
    }
}
