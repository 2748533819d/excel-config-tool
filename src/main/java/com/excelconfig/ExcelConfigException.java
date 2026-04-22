package com.excelconfig;

/**
 * Excel 配置工具异常类
 */
public class ExcelConfigException extends RuntimeException {

    public ExcelConfigException(String message) {
        super(message);
    }

    public ExcelConfigException(String message, Throwable cause) {
        super(message, cause);
    }
}
