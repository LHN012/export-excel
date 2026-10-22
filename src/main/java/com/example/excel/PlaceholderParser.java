package com.example.excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** 只识别占位符，不负责查数据或操作 Excel。 */
final class PlaceholderParser {
    private static final Pattern PLACEHOLDER = Pattern.compile("\\$\\{([^{}]*)}");

    Matcher matcher(String text) {
        return PLACEHOLDER.matcher(text);
    }

    String path(Matcher matcher) {
        String path = matcher.group(1).trim();
        // 点号分隔对象属性；* 无法确定目标列，模板需要明确写出字段名。
        for (String part : path.split("\\.", -1)) {
            if (part.isEmpty() || part.contains("*") || !part.equals(part.trim())) {
                throw new ExcelExportException("无效占位符 " + matcher.group()
                        + "，请使用 ${name} 或 ${orders.productName} 这样的字段路径");
            }
        }
        return path;
    }
}
