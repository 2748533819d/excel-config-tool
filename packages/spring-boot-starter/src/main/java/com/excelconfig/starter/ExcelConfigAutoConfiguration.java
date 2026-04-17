package com.excelconfig.starter;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.extract.ExtractEngine;
import com.excelconfig.export.FillEngine;
import com.excelconfig.locator.HeaderLocator;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

/**
 * Excel Config 自动配置
 *
 * 自动配置以下 Bean：
 * - JsonConfigParser: JSON 配置解析器
 * - HeaderLocator: 表头定位器
 * - ExtractEngine: 数据提取引擎
 * - FillEngine: 数据填充引擎
 * - ExcelConfigService: Excel 配置服务
 */
@AutoConfiguration
@ConditionalOnClass({JsonConfigParser.class, ExtractEngine.class, FillEngine.class})
@EnableConfigurationProperties(ExcelConfigProperties.class)
@ConditionalOnProperty(prefix = "excel.config", name = "enabled", havingValue = "true", matchIfMissing = true)
public class ExcelConfigAutoConfiguration {

    @Bean
    @ConditionalOnMissingBean
    public JsonConfigParser jsonConfigParser() {
        return new JsonConfigParser();
    }

    @Bean
    @ConditionalOnMissingBean
    public HeaderLocator headerLocator() {
        return new HeaderLocator();
    }

    @Bean
    @ConditionalOnMissingBean
    public ExtractEngine extractEngine() {
        return new ExtractEngine();
    }

    @Bean
    @ConditionalOnMissingBean
    public FillEngine fillEngine() {
        return new FillEngine();
    }

    @Bean
    @ConditionalOnMissingBean
    public ExcelConfigService excelConfigService(
            JsonConfigParser jsonConfigParser,
            ExtractEngine extractEngine,
            FillEngine fillEngine,
            ExcelConfigProperties properties) {
        return new ExcelConfigService(jsonConfigParser, extractEngine, fillEngine, properties);
    }
}
