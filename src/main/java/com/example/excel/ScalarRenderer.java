package com.example.excel;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.regex.Matcher;

/** 单元格替换：纯占位符保留数字/布尔类型，混合文本输出字符串。 */
final class ScalarRenderer {
    private final PlaceholderParser parser;
    private final ValueResolver resolver;

    ScalarRenderer(PlaceholderParser parser, ValueResolver resolver) {
        this.parser = parser;
        this.resolver = resolver;
    }

    void render(Row row, JsonNode root, String listPath, JsonNode item) {
        for (Cell cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                try {
                    render(cell, root, listPath, item);
                } catch (ExcelExportException e) {
                    throw new ExcelExportException(location(cell) + e.getMessage(), e);
                }
            }
        }
    }

    private void render(Cell cell, JsonNode root, String listPath, JsonNode item) {
        String template = cell.getStringCellValue();
        Matcher matcher = parser.matcher(template);
        if (matcher.matches()) {
            JsonNode value = value(root, matcher, listPath, item);
            if (value.isNull() || value.isMissingNode()) {
                cell.setBlank();
            } else if (value.isBoolean()) {
                cell.setCellValue(value.booleanValue());
            } else if (isExcelNumber(value)) {
                cell.setCellValue(value.doubleValue());
            } else {
                // Excel 数字只有约 15 位有效精度，长编号或高精度小数改用文本保存。
                cell.setCellValue(value.asText());
            }
            return;
        }
        StringBuffer result = new StringBuffer();
        boolean replaced = false;
        while (matcher.find()) {
            JsonNode value = value(root, matcher, listPath, item);
            String replacement = value.isNull() || value.isMissingNode() ? "" : value.asText();
            // JSON 值可能含有 $ 和反斜杠，必须转义正则替换串。
            matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
            replaced = true;
        }
        if (replaced) {
            matcher.appendTail(result);
            cell.setCellValue(result.toString());
        }
    }

    private boolean isExcelNumber(JsonNode value) {
        if (!value.isNumber() || value.decimalValue().precision() > 15) {
            return false;
        }
        double number = value.doubleValue();
        // 超出 Excel 范围时按文本保存，避免溢出或下溢后静默变成 0。
        return Double.isFinite(number) && (number == 0
                ? value.decimalValue().signum() == 0 : Math.abs(number) >= Double.MIN_NORMAL);
    }

    private JsonNode value(JsonNode root, Matcher matcher, String listPath, JsonNode item) {
        String path = parser.path(matcher);
        JsonNode value = resolver.resolve(root, path, listPath, item);
        if (value.isContainerNode()) {
            throw new ExcelExportException("${" + path
                    + "} 对应对象或嵌套数组，请指定一个普通字段");
        }
        return value;
    }

    static String location(Cell cell) {
        return "工作表「" + cell.getSheet().getSheetName() + "」单元格 "
                + cell.getAddress().formatAsString() + "：";
    }
}
