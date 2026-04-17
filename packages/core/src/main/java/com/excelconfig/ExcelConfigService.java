package com.excelconfig;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.export.FillEngine;
import com.excelconfig.extract.ExtractEngine;
import com.excelconfig.model.ExcelConfig;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

/**
 * Excel 配置工具门面接口
 *
 * 提供简化的统一入口，用于 Excel 数据的提取和填充操作
 */
public class ExcelConfigService {

    private final ExtractEngine extractEngine;
    private final FillEngine fillEngine;
    private final JsonConfigParser configParser;

    public ExcelConfigService() {
        this.extractEngine = new ExtractEngine();
        this.fillEngine = new FillEngine();
        this.configParser = new JsonConfigParser();
    }

    /**
     * 从 Excel 文件中提取数据
     *
     * @param template Excel 模板输入流
     * @param configJson 配置 JSON 字符串
     * @return 提取的数据 Map
     */
    public Map<String, Object> extract(InputStream template, String configJson) {
        ExcelConfig config = parseConfig(configJson);
        return extractEngine.extract(template, config);
    }

    /**
     * 从 Excel 文件中提取数据
     *
     * @param template Excel 模板输入流
     * @param config 配置对象
     * @return 提取的数据 Map
     */
    public Map<String, Object> extract(InputStream template, ExcelConfig config) {
        return extractEngine.extract(template, config);
    }

    /**
     * 填充数据到 Excel 文件
     *
     * @param template Excel 模板输入流
     * @param data 数据 Map
     * @param configJson 配置 JSON 字符串
     * @return 填充后的 Excel 文件字节数组
     */
    public byte[] fill(InputStream template, Map<String, Object> data, String configJson) {
        ExcelConfig config = parseConfig(configJson);
        return fillEngine.fill(template, data, config);
    }

    /**
     * 填充数据到 Excel 文件
     *
     * @param template Excel 模板输入流
     * @param data 数据 Map
     * @param config 配置对象
     * @return 填充后的 Excel 文件字节数组
     */
    public byte[] fill(InputStream template, Map<String, Object> data, ExcelConfig config) {
        return fillEngine.fill(template, data, config);
    }

    /**
     * 填充数据到 Excel 文件并输出到指定流
     *
     * @param template Excel 模板输入流
     * @param data 数据 Map
     * @param configJson 配置 JSON 字符串
     * @param output 输出流
     */
    public void fill(InputStream template, Map<String, Object> data, String configJson, OutputStream output) {
        byte[] result = fill(template, data, configJson);
        try {
            output.write(result);
            output.flush();
        } catch (Exception e) {
            throw new ExcelConfigException("写入输出流失败", e);
        }
    }

    /**
     * 填充数据到 Excel 文件并输出到指定流
     *
     * @param template Excel 模板输入流
     * @param data 数据 Map
     * @param config 配置对象
     * @param output 输出流
     */
    public void fill(InputStream template, Map<String, Object> data, ExcelConfig config, OutputStream output) {
        byte[] result = fill(template, data, config);
        try {
            output.write(result);
            output.flush();
        } catch (Exception e) {
            throw new ExcelConfigException("写入输出流失败", e);
        }
    }

    /**
     * 解析配置 JSON 字符串
     *
     * @param configJson 配置 JSON 字符串
     * @return 配置对象
     */
    public ExcelConfig parseConfig(String configJson) {
        try {
            return configParser.parse(configJson);
        } catch (Exception e) {
            throw new ExcelConfigException("解析配置失败：" + e.getMessage(), e);
        }
    }
}
