package com.bonitasoft.connectors.excel;

/**
 * Typed exception for Excel connector operations.
 */
public class ExcelException extends Exception {

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }
}
