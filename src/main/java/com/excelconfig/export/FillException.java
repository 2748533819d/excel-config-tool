package com.excelconfig.export;

/**
 * 填充异常
 */
public class FillException extends RuntimeException {

    public FillException(String message) {
        super(message);
    }

    public FillException(String message, Throwable cause) {
        super(message, cause);
    }
}
