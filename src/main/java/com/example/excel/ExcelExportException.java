package com.example.excel;

/** JSON 数据或模板占位符不符合导出规则时抛出。 */
public class ExcelExportException extends IllegalArgumentException {
    private static final long serialVersionUID = 1L;

    public ExcelExportException(String message) {
        super(message);
    }

    public ExcelExportException(String message, Throwable cause) {
        super(message, cause);
    }
}
