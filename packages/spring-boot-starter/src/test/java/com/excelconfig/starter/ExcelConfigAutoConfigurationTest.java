package com.excelconfig.starter;

import org.junit.jupiter.api.Test;
import org.springframework.boot.autoconfigure.AutoConfigurations;
import org.springframework.boot.test.context.runner.ApplicationContextRunner;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Excel Config 自动配置测试
 */
class ExcelConfigAutoConfigurationTest {

    private final ApplicationContextRunner contextRunner = new ApplicationContextRunner()
            .withConfiguration(AutoConfigurations.of(ExcelConfigAutoConfiguration.class));

    @Test
    void testDefaultInitialization() {
        contextRunner.run(context -> {
            assertTrue(context.containsBean("jsonConfigParser"));
            assertTrue(context.containsBean("headerLocator"));
            assertTrue(context.containsBean("extractEngine"));
            assertTrue(context.containsBean("fillEngine"));
            assertTrue(context.containsBean("excelConfigService"));
        });
    }

    @Test
    void testDisabled() {
        contextRunner
            .withPropertyValues("excel.config.enabled=false")
            .run(context -> {
                assertFalse(context.containsBean("jsonConfigParser"));
                assertFalse(context.containsBean("extractEngine"));
                assertFalse(context.containsBean("fillEngine"));
                assertFalse(context.containsBean("excelConfigService"));
            });
    }

    @Test
    void testCustomProperties() {
        contextRunner
            .withPropertyValues(
                "excel.config.template-location=classpath:templates/",
                "excel.config.output-location=classpath:output/"
            )
            .run(context -> {
                ExcelConfigProperties properties = context.getBean(ExcelConfigProperties.class);
                assertEquals("classpath:templates/", properties.getTemplateLocation());
                assertEquals("classpath:output/", properties.getOutputLocation());
            });
    }
}
