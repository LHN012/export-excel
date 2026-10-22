package com.example.excel;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.MissingNode;

/** 支持 customer.name 这样的对象路径，并在列表行中切换到当前数组元素。 */
final class ValueResolver {

    /** 返回路径中遇到的第一个数组的完整路径，没有数组时返回 null。 */
    String findListPath(JsonNode root, String path) {
        JsonNode current = root;
        StringBuilder visited = new StringBuilder();
        for (String part : path.split("\\.")) {
            if (current.isArray()) {
                return visited.toString();
            }
            current = child(current, part);
            if (visited.length() > 0) {
                visited.append('.');
            }
            visited.append(part);
        }
        return current.isArray() ? visited.toString() : null;
    }

    JsonNode resolve(JsonNode root, String path) {
        return resolve(root, path, null, null);
    }

    JsonNode resolve(JsonNode root, String path, String listPath, JsonNode item) {
        JsonNode current = root;
        String remaining = path;
        if (listPath != null) {
            if (path.equals(listPath)) {
                return item;
            }
            if (path.startsWith(listPath + ".")) {
                current = item;
                remaining = path.substring(listPath.length() + 1);
            }
        }
        for (String part : remaining.split("\\.")) {
            if (current.isArray()) {
                throw new ExcelExportException("路径 " + path
                        + " 包含未展开的数组；当前版本不支持嵌套列表");
            }
            current = child(current, part);
        }
        return current;
    }

    private JsonNode child(JsonNode current, String field) {
        JsonNode value = current.isObject() ? current.get(field) : null;
        return value == null ? MissingNode.getInstance() : value;
    }
}
