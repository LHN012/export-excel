package com.example.excel;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;

/**
 * 将一个 JSON 对象填入 XLSX 模板的全部工作表，无需 Spring。
 * <p>支持 ${customer.name} 和 ${orders.productName} 这样的对象/列表路径。
 * 每个列表模板行只能引用一个数组，空数组会删除该行。流由调用方负责关闭。
 * 同一个导出器可以重复使用；每次调用都创建独立的工作簿。</p>
 */
public class ExcelExporter {
    private final ObjectMapper mapper = new ObjectMapper()
            .enable(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS)
            .enable(DeserializationFeature.FAIL_ON_TRAILING_TOKENS)
            .enable(JsonParser.Feature.STRICT_DUPLICATE_DETECTION);
    private final PlaceholderParser parser = new PlaceholderParser();
    private final ValueResolver resolver = new ValueResolver();
    private final ScalarRenderer scalarRenderer = new ScalarRenderer(parser, resolver);
    private final ListRenderer listRenderer = new ListRenderer(scalarRenderer);

    /**
     * @param template XLSX 模板输入流
     * @param output 导出结果输出流，不会被本方法关闭
     * @param json 根节点为对象的 JSON 字符串
     * @throws IOException JSON 解析失败、模板读取失败或结果写入失败
     * @throws ExcelExportException 模板或数据不符合占位符规则
     */
    public void export(InputStream template, OutputStream output, String json) throws IOException {
        Objects.requireNonNull(template, "template 不能为空");
        Objects.requireNonNull(output, "output 不能为空");
        Objects.requireNonNull(json, "json 不能为空");
        JsonNode root = mapper.readTree(json);
        if (root == null || !root.isObject()) {
            throw new ExcelExportException("JSON 根节点必须是对象，例如 {\"name\":\"张三\"}");
        }

        // Workbook 关闭时可能关闭底层输入流，用包装器保证调用方拥有流的生命周期。
        InputStream borrowed = new FilterInputStream(template) {
            @Override
            public void close() {
                // 不关闭调用方传入的流。
            }
        };
        try (XSSFWorkbook workbook = new XSSFWorkbook(borrowed)) {
            List<List<RowPlan>> plans = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                plans.add(plan(workbook.getSheetAt(i), root));
            }
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                for (RowPlan row : plans.get(i)) {
                    if (row.listPath == null) {
                        scalarRenderer.render(sheet.getRow(row.index), root, null, null);
                    } else {
                        listRenderer.render(sheet, row.index, root, row.listPath, row.list);
                    }
                }
            }
            // 让 Excel 打开文件时重新计算公式，避免沿用模板中的缓存结果。
            workbook.setForceFormulaRecalculation(true);
            workbook.write(output);
        }
    }

    private List<RowPlan> plan(XSSFSheet sheet, JsonNode root) {
        List<RowPlan> rows = new ArrayList<>();
        for (Row row : sheet) {
            String listPath = null;
            boolean hasPlaceholder = false;
            for (Cell cell : row) {
                if (cell.getCellType() != CellType.STRING) {
                    continue;
                }
                Matcher matcher = parser.matcher(cell.getStringCellValue());
                while (matcher.find()) {
                    hasPlaceholder = true;
                    String candidate;
                    try {
                        candidate = resolver.findListPath(root, parser.path(matcher));
                    } catch (ExcelExportException e) {
                        throw new ExcelExportException(ScalarRenderer.location(cell) + e.getMessage(), e);
                    }
                    if (candidate != null) {
                        if (listPath != null && !listPath.equals(candidate)) {
                            throw new ExcelExportException(ScalarRenderer.location(cell)
                                    + "同一模板行引用了不同数组：" + listPath + " 和 " + candidate);
                        }
                        listPath = candidate;
                    }
                }
            }
            if (hasPlaceholder) {
                if (listPath != null) {
                    listRenderer.validate(sheet, row.getRowNum());
                }
                rows.add(new RowPlan(row.getRowNum(), listPath,
                        listPath == null ? null : resolver.resolve(root, listPath)));
            }
        }
        // 原始行号从下向上处理。下方插入/删除行，不会改变尚未处理的上方行号。
        // 每行只替换一次，JSON 值里出现的 ${...} 不会被再次解析。
        Collections.reverse(rows);
        return rows;
    }

    private static final class RowPlan {
        private final int index;
        private final String listPath;
        private final JsonNode list;

        private RowPlan(int index, String listPath, JsonNode list) {
            this.index = index;
            this.listPath = listPath;
            this.list = list;
        }
    }
}
