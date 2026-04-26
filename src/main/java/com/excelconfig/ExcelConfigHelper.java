package com.excelconfig;

import com.excelconfig.config.JsonConfigParser;
import com.excelconfig.export.FillEngine;
import com.excelconfig.extract.ExtractEngine;
import com.excelconfig.model.ExcelConfig;
import com.excelconfig.util.JsonUtil;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

/**
 * Excel 配置工具门面类 - 统一入口（类似 EasyExcel 风格）
 *
 * 提供简洁的静态方法用于 Excel 数据的提取和填充操作
 *
 * 使用示例：
 * <pre>
 * // 提取数据
 * Map&lt;String, Object&gt; data = ExcelConfigHelper.read("template.xlsx")
 *     .config("config.json")
 *     .extract();
 *
 * // 填充数据
 * ExcelConfigHelper.write("template.xlsx")
 *     .config("config.json")
 *     .data(data)
 *     .writeTo("output.xlsx");
 * </pre>
 */
public class ExcelConfigHelper {

    private final ExtractEngine extractEngine;
    private final FillEngine fillEngine;
    private final JsonConfigParser configParser;

    private String configJson;
    private ExcelConfig configObject;
    private InputStream templateStream;
    private File templateFile;
    private Path templatePath;
    private Map<String, Object> data;

    public ExcelConfigHelper() {
        this.extractEngine = new ExtractEngine();
        this.fillEngine = new FillEngine();
        this.configParser = new JsonConfigParser();
    }

    /**
     * 创建读取器（用于提取数据）
     *
     * @param templateFile Excel 模板文件
     * @return ExcelConfigHelper 构建器
     */
    public static ExcelConfigHelper read(File templateFile) {
        return new ExcelConfigHelper().setTemplateFile(templateFile);
    }

    /**
     * 创建读取器（用于提取数据）
     *
     * @param templatePath Excel 模板文件路径
     * @return ExcelConfigHelper 构建器
     */
    public static ExcelConfigHelper read(String templatePath) {
        return new ExcelConfigHelper().setTemplatePath(Path.of(templatePath));
    }

    /**
     * 创建读取器（用于提取数据）
     *
     * @param templateStream Excel 模板输入流
     * @return ExcelConfigHelper 构建器
     */
    public static ExcelConfigHelper read(InputStream templateStream) {
        return new ExcelConfigHelper().setTemplateStream(templateStream);
    }

    /**
     * 创建写入器（用于填充数据）
     *
     * @param templateFile Excel 模板文件
     * @return ExcelConfigHelper 构建器
     */
    public static ExcelConfigHelper write(File templateFile) {
        return new ExcelConfigHelper().setTemplateFile(templateFile);
    }

    /**
     * 创建写入器（用于填充数据）
     *
     * @param templatePath Excel 模板文件路径
     * @return ExcelConfigHelper 构建器
     */
    public static ExcelConfigHelper write(String templatePath) {
        return new ExcelConfigHelper().setTemplatePath(Path.of(templatePath));
    }

    /**
     * 创建写入器（用于填充数据）
     *
     * @param templateStream Excel 模板输入流
     * @return ExcelConfigHelper 构建器
     */
    public static ExcelConfigHelper write(InputStream templateStream) {
        return new ExcelConfigHelper().setTemplateStream(templateStream);
    }

    // ========== 配置方法 ==========

    /**
     * 设置配置文件路径
     *
     * @param configPath 配置文件路径
     * @return this
     */
    public ExcelConfigHelper config(String configPath) throws Exception {
        Path path = Path.of(configPath);
        this.configJson = Files.readString(path);
        return this;
    }

    /**
     * 设置配置文件
     *
     * @param configFile 配置文件
     * @return this
     */
    public ExcelConfigHelper config(File configFile) throws Exception {
        this.configJson = Files.readString(configFile.toPath());
        return this;
    }

    /**
     * 设置配置 JSON 字符串
     *
     * @param configJson 配置 JSON 字符串
     * @return this
     */
    public ExcelConfigHelper configJson(String configJson) {
        this.configJson = configJson;
        return this;
    }

    /**
     * 设置配置对象
     *
     * @param config 配置对象
     * @return this
     */
    public ExcelConfigHelper configObject(ExcelConfig config) {
        this.configObject = config;
        return this;
    }

    // ========== 数据方法 ==========

    /**
     * 设置要填充的数据
     *
     * @param data 数据 Map
     * @return this
     */
    public ExcelConfigHelper data(Map<String, Object> data) {
        this.data = data;
        return this;
    }

    // ========== 提取方法 ==========

    /**
     * 执行提取操作
     *
     * @return 提取的数据 Map
     */
    public Map<String, Object> extract() {
        validateExtractParams();
        try {
            ExcelConfig config = getConfigObject();
            return extractEngine.extract(getTemplateStream(), config);
        } catch (Exception e) {
            throw new ExcelConfigException("提取数据失败：" + e.getMessage(), e);
        }
    }

    /**
     * 执行提取操作并转换为指定类型
     *
     * @param clazz 目标类型
     * @param <T> 类型参数
     * @return 转换后的对象
     */
    public <T> T extractAs(Class<T> clazz) {
        Map<String, Object> result = extract();
        return convertMapToObject(result, clazz);
    }

    // ========== 填充方法 ==========

    /**
     * 执行填充操作，返回字节数组
     *
     * @return 填充后的 Excel 文件字节数组
     */
    public byte[] write() {
        validateFillParams();
        try {
            ExcelConfig config = getConfigObject();
            return fillEngine.fill(getTemplateStream(), data, config);
        } catch (Exception e) {
            throw new ExcelConfigException("填充数据失败：" + e.getMessage(), e);
        }
    }

    /**
     * 执行填充操作并写入文件
     *
     * @param outputPath 输出文件路径
     */
    public void writeTo(String outputPath) {
        writeTo(Path.of(outputPath));
    }

    /**
     * 执行填充操作并写入文件
     *
     * @param outputFile 输出文件
     */
    public void writeTo(File outputFile) {
        writeTo(outputFile.toPath());
    }

    /**
     * 执行填充操作并写入文件
     *
     * @param outputPath 输出文件路径
     */
    public void writeTo(Path outputPath) {
        byte[] result = write();
        try {
            Files.write(outputPath, result);
        } catch (Exception e) {
            throw new ExcelConfigException("写入文件失败：" + e.getMessage(), e);
        }
    }

    /**
     * 执行填充操作并写入输出流
     *
     * @param output 输出流
     */
    public void writeTo(OutputStream output) {
        byte[] result = write();
        try {
            output.write(result);
            output.flush();
        } catch (Exception e) {
            throw new ExcelConfigException("写入输出流失败：" + e.getMessage(), e);
        }
    }

    // ========== 私有方法 ==========

    private ExcelConfig getConfigObject() throws Exception {
        if (configObject != null) {
            return configObject;
        }
        if (configJson != null) {
            return configParser.parse(configJson);
        }
        throw new ExcelConfigException("必须指定配置文件或配置 JSON");
    }

    private InputStream getTemplateStream() {
        try {
            if (templateStream != null) {
                return templateStream;
            }
            if (templateFile != null) {
                return Files.newInputStream(templateFile.toPath());
            }
            if (templatePath != null) {
                return Files.newInputStream(templatePath);
            }
        } catch (Exception e) {
            throw new ExcelConfigException("打开模板文件失败：" + e.getMessage(), e);
        }
        throw new ExcelConfigException("必须指定 Excel 模板文件");
    }

    private void validateExtractParams() {
        if (configJson == null && configObject == null) {
            throw new ExcelConfigException("必须指定配置文件或配置 JSON");
        }
    }

    private void validateFillParams() {
        if (configJson == null && configObject == null) {
            throw new ExcelConfigException("必须指定配置文件或配置 JSON");
        }
        if (data == null) {
            throw new ExcelConfigException("必须指定要填充的数据");
        }
    }

    private <T> T convertMapToObject(Map<String, Object> map, Class<T> clazz) {
        return JsonUtil.convertToObject(map, clazz);
    }

    // ========== Setter 方法 ==========

    private ExcelConfigHelper setTemplateFile(File templateFile) {
        this.templateFile = templateFile;
        return this;
    }

    private ExcelConfigHelper setTemplatePath(Path templatePath) {
        this.templatePath = templatePath;
        return this;
    }

    private ExcelConfigHelper setTemplateStream(InputStream templateStream) {
        this.templateStream = templateStream;
        return this;
    }
}
