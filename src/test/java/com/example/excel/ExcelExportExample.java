package com.example.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/** 在 IDEA 中运行 main：首次生成示例模板，随后使用 JSON 导出 Excel。 */
public class ExcelExportExample {
    public static void main(String[] args) throws Exception {
        Path project = args.length == 0 ? Paths.get("") : Paths.get(args[0]);
        Path template = project.resolve("examples/template.xlsx");
        Path data = project.resolve("examples/data.json");
        Path output = project.resolve("target/example-output.xlsx");
        Files.createDirectories(template.getParent());
        Files.createDirectories(output.getParent());

        // 已有模板可直接修改；重复运行示例不会覆盖用户编辑过的模板。
        if (!Files.exists(template)) {
            createTemplate(template);
        }
        String json = new String(Files.readAllBytes(data), StandardCharsets.UTF_8);
        try (InputStream input = Files.newInputStream(template);
             OutputStream result = Files.newOutputStream(output)) {
            new ExcelExporter().export(input, result, json);
        }
        System.out.println("模板：" + template.toAbsolutePath());
        System.out.println("JSON：" + data.toAbsolutePath());
        System.out.println("结果：" + output.toAbsolutePath());
    }

    private static void createTemplate(Path path) throws Exception {
        try (XSSFWorkbook book = new XSSFWorkbook()) {
            XSSFCellStyle header = headerStyle(book);
            XSSFCellStyle normal = normalStyle(book);
            XSSFCellStyle money = book.createCellStyle();
            money.cloneStyleFrom(normal);
            money.setDataFormat(book.createDataFormat().getFormat("0.00"));

            XSSFSheet orders = book.createSheet("订单明细");
            setupColumns(orders);
            title(orders, 0, "订单明细", header);
            row(orders, 1, normal, "客户：${customer.name}", "电话：${customer.phone}");
            row(orders, 2, normal, "导出日期：${exportedAt}");
            row(orders, 4, header, "商品", "数量", "单价", "金额");
            XSSFRow order = row(orders, 5, normal,
                    "${orders.productName}", "${orders.quantity}", "${orders.unitPrice}");
            order.getCell(2).setCellStyle(money);
            XSSFCell amount = order.createCell(3);
            amount.setCellFormula("B6*C6");
            amount.setCellStyle(money);
            row(orders, 6, normal, "订单合计", "${orderTotal}");

            XSSFSheet payments = book.createSheet("收款与备注");
            setupColumns(payments);
            title(payments, 0, "收款明细", header);
            row(payments, 1, normal, "客户：${customer.name}");
            row(payments, 3, header, "收款日期", "金额");
            XSSFRow payment = row(payments, 4, normal, "${payments.date}", "${payments.amount}");
            payment.getCell(1).setCellStyle(money);
            title(payments, 6, "备注", header);
            row(payments, 7, normal, "${tags}");
            row(payments, 8, normal, "导出完成");

            book.setForceFormulaRecalculation(true);
            try (OutputStream output = Files.newOutputStream(path)) {
                book.write(output);
            }
        }
    }

    private static void setupColumns(XSSFSheet sheet) {
        sheet.setColumnWidth(0, 36 * 256);
        sheet.setColumnWidth(1, 32 * 256);
        sheet.setColumnWidth(2, 26 * 256);
        sheet.setColumnWidth(3, 20 * 256);
        sheet.createFreezePane(0, 1);
    }

    private static void title(XSSFSheet sheet, int index, String text, XSSFCellStyle style) {
        row(sheet, index, style, text, "", "", "");
        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 3));
    }

    private static XSSFRow row(XSSFSheet sheet, int index, XSSFCellStyle style, String... values) {
        XSSFRow row = sheet.createRow(index);
        row.setHeightInPoints(26);
        for (int i = 0; i < values.length; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(values[i]);
            cell.setCellStyle(style);
        }
        return row;
    }

    private static XSSFCellStyle normalStyle(XSSFWorkbook book) {
        XSSFFont font = book.createFont();
        font.setFontName("Microsoft YaHei");
        font.setFontHeightInPoints((short) 11);
        XSSFCellStyle style = book.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static XSSFCellStyle headerStyle(XSSFWorkbook book) {
        XSSFCellStyle style = normalStyle(book);
        XSSFFont font = book.createFont();
        font.setFontName("Microsoft YaHei");
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }
}
