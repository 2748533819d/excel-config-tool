package com.excelconfig.starter;

import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * Excel Config 配置属性
 *
 * 使用示例：
 * <pre>
 * excel.config:
 *   enabled: true
 *   template-location: classpath:templates/
 *   output-location: classpath:output/
 * </pre>
 */
@ConfigurationProperties(prefix = "excel.config")
public class ExcelConfigProperties {

    /**
     * 是否启用 Excel Config 功能
     */
    private boolean enabled = true;

    /**
     * 模板文件位置
     */
    private String templateLocation = "classpath:templates/";

    /**
     * 输出文件位置
     */
    private String outputLocation = "classpath:output/";

    public boolean isEnabled() {
        return enabled;
    }

    public void setEnabled(boolean enabled) {
        this.enabled = enabled;
    }

    public String getTemplateLocation() {
        return templateLocation;
    }

    public void setTemplateLocation(String templateLocation) {
        this.templateLocation = templateLocation;
    }

    public String getOutputLocation() {
        return outputLocation;
    }

    public void setOutputLocation(String outputLocation) {
        this.outputLocation = outputLocation;
    }
}
