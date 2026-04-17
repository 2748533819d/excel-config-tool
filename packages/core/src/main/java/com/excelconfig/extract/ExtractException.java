package com.excelconfig.extract;

/**
 * 提取异常
 */
public class ExtractException extends RuntimeException {

    public ExtractException(String message) {
        super(message);
    }

    public ExtractException(String message, Throwable cause) {
        super(message, cause);
    }
}
