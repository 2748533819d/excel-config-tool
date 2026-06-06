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

    /**
     * 是否启用 SXSSF 流式写入（大幅降低内存占用，适合大数据量导出）
     */
    private Boolean streaming;

    /**
     * SXSSF 行缓存窗口大小（默认 100），超过此值的行将被刷入磁盘
     */
    private Integer streamingRowWindowSize;

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

    public Boolean getStreaming() {
        return streaming;
    }

    public void setStreaming(Boolean streaming) {
        this.streaming = streaming;
    }

    public Integer getStreamingRowWindowSize() {
        return streamingRowWindowSize;
    }

    public void setStreamingRowWindowSize(Integer streamingRowWindowSize) {
        this.streamingRowWindowSize = streamingRowWindowSize;
    }
}
