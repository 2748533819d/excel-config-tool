package com.excelconfig.locator;

/**
 * 表头未找到异常
 */
public class HeaderNotFoundException extends RuntimeException {

    public HeaderNotFoundException(String message) {
        super(message);
    }

    public HeaderNotFoundException(String message, Throwable cause) {
        super(message, cause);
    }
}
