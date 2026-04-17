package com.excelconfig.starter;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.extract.ExtractEngine;
import com.excelconfig.export.FillEngine;
import com.excelconfig.model.ExcelConfig;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.util.StreamUtils;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Map;

/**
 * Excel Config 服务
 *
 * 提供简化的 API 用于：
 * - 加载和解析 Excel 配置文件
 * - 执行数据提取
 * - 执行数据填充
 */
public class ExcelConfigService {

    private final JsonConfigParser configParser;
    private final ExtractEngine extractEngine;
    private final FillEngine fillEngine;
    private final ExcelConfigProperties properties;

    public ExcelConfigService(JsonConfigParser configParser,
                              ExtractEngine extractEngine,
                              FillEngine fillEngine,
                              ExcelConfigProperties properties) {
        this.configParser = configParser;
        this.extractEngine = extractEngine;
        this.fillEngine = fillEngine;
        this.properties = properties;
    }

    /**
     * 从 classpath 加载配置
     *
     * @param location 配置文件位置，如 classpath:config/excel-config.json
     * @return ExcelConfig 配置对象
     * @throws IOException 读取文件失败
     */
    public ExcelConfig loadConfig(String location) throws IOException {
        if (location.startsWith("classpath:")) {
            String path = location.substring("classpath:".length());
            Resource resource = new org.springframework.core.io.ClassPathResource(path);
            try (InputStream input = resource.getInputStream()) {
                String json = StreamUtils.copyToString(input, StandardCharsets.UTF_8);
                return configParser.parse(json);
            }
        } else if (location.startsWith("file:")) {
            String path = location.substring("file:".length());
            Resource resource = new org.springframework.core.io.FileSystemResource(path);
            try (InputStream input = resource.getInputStream()) {
                String json = StreamUtils.copyToString(input, StandardCharsets.UTF_8);
                return configParser.parse(json);
            }
        } else {
            throw new IllegalArgumentException("不支持的资源位置：" + location);
        }
    }

    /**
     * 从 JSON 字符串解析配置
     *
     * @param json JSON 配置字符串
     * @return ExcelConfig 配置对象
     * @throws IOException 解析失败
     */
    public ExcelConfig parseConfig(String json) throws IOException {
        return configParser.parse(json);
    }

    /**
     * 执行数据提取
     *
     * @param config Excel 配置
     * @param templateStream Excel 模板输入流
     * @return 提取的数据
     */
    public Map<String, Object> extract(ExcelConfig config, InputStream templateStream) {
        return extractEngine.extract(templateStream, config);
    }

    /**
     * 执行数据填充
     *
     * @param config Excel 配置
     * @param data 数据
     * @param templateStream Excel 模板输入流
     * @return 填充后的 Excel 文件字节数组
     */
    public byte[] fill(ExcelConfig config, Map<String, Object> data, InputStream templateStream) {
        return fillEngine.fill(templateStream, data, config);
    }

    /**
     * 获取配置属性
     *
     * @return ExcelConfigProperties
     */
    public ExcelConfigProperties getProperties() {
        return properties;
    }
}
