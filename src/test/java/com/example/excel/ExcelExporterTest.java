package com.example.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/** 导出后重新打开 XLSX，验证实际文件内容。 */
class ExcelExporterTest {
    private final ExcelExporter exporter = new ExcelExporter();

    @Test
    void replacesScalarPathsAndPreservesCellTypes() throws Exception {
        byte[] template = template(book -> {
            XSSFRow row = book.createSheet("客户").createRow(0);
            cells(row, "${customer.name}", "客户：${customer.name} / ${customer.phone}",
                    "${quantity}", "${active}", "${missing}", "${empty}", "原有文本", "${money}");
            row.createCell(8).setCellValue(99);
            row.createCell(9).setCellFormula("C1*2");
        });
        try (XSSFWorkbook result = export(template, "{\"customer\":{\"name\":\"张三\",\"phone\":\"138\"},"
                + "\"quantity\":2,\"active\":true,\"empty\":null,\"money\":12.5}")) {
            XSSFRow row = result.getSheetAt(0).getRow(0);
            assertEquals("张三", row.getCell(0).getStringCellValue());
            assertEquals("客户：张三 / 138", row.getCell(1).getStringCellValue());
            assertEquals(CellType.NUMERIC, row.getCell(2).getCellType());
            assertEquals(2, row.getCell(2).getNumericCellValue());
            assertTrue(row.getCell(3).getBooleanCellValue());
            assertEquals(CellType.BLANK, row.getCell(4).getCellType());
            assertEquals(CellType.BLANK, row.getCell(5).getCellType());
            assertEquals("原有文本", row.getCell(6).getStringCellValue());
            assertEquals(12.5, row.getCell(7).getNumericCellValue());
            assertEquals(99, row.getCell(8).getNumericCellValue());
            assertEquals("C1*2", row.getCell(9).getCellFormula());
        }
    }

    @Test
    void expandsIndependentListsAcrossMultipleSheetsAndKeepsFooters() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet orders = book.createSheet("订单");
            cells(orders.createRow(0), "客户：${customer.name}");
            cells(orders.createRow(2), "${orders.name}", "${orders.quantity}", "${customer.name}");
            cells(orders.createRow(3), "订单结束");
            cells(orders.createRow(5), "${payments.date}", "${payments.amount}");
            cells(orders.createRow(6), "所有数据结束：${customer.name}");
            XSSFSheet payments = book.createSheet("收款");
            cells(payments.createRow(0), "${payments.date}", "${payments.amount}");
            cells(payments.createRow(1), "收款结束");
            cells(book.createSheet("重复使用订单").createRow(0), "${orders.name}");
        });
        String json = "{\"customer\":{\"name\":\"张三\"},"
                + "\"orders\":[{\"name\":\"钢笔\",\"quantity\":2},{\"name\":\"纸\",\"quantity\":5}],"
                + "\"payments\":[{\"date\":\"周一\",\"amount\":10},{\"date\":\"周二\",\"amount\":20},"
                + "{\"date\":\"周三\",\"amount\":30}]}";
        try (XSSFWorkbook result = export(template, json)) {
            XSSFSheet sheet = result.getSheetAt(0);
            assertEquals("客户：张三", text(sheet, 0, 0));
            assertEquals("钢笔", text(sheet, 2, 0));
            assertEquals("纸", text(sheet, 3, 0));
            assertEquals(5, sheet.getRow(3).getCell(1).getNumericCellValue());
            assertEquals("张三", text(sheet, 3, 2));
            assertEquals("订单结束", text(sheet, 4, 0));
            assertEquals("周一", text(sheet, 6, 0));
            assertEquals("周三", text(sheet, 8, 0));
            assertEquals("所有数据结束：张三", text(sheet, 9, 0));
            assertEquals("周三", text(result.getSheetAt(1), 2, 0));
            assertEquals("收款结束", text(result.getSheetAt(1), 3, 0));
            assertEquals("纸", text(result.getSheetAt(2), 1, 0));
        }
    }

    @Test
    void copiesStylesHeightHorizontalMergesAndRelativeFormulas() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet sheet = book.createSheet("明细");
            sheet.setColumnWidth(0, 6000);
            XSSFRow row = sheet.createRow(1);
            row.setHeightInPoints(32);
            cells(row, "${orders.name}", "", "${orders.quantity}", "${orders.price}");
            row.createCell(4).setCellFormula("C2*D2+$H$1");
            XSSFCellStyle style = book.createCellStyle();
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.THIN);
            style.setDataFormat(book.createDataFormat().getFormat("0.00"));
            row.getCell(3).setCellStyle(style);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            cells(sheet.createRow(2), "尾部");
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
        });
        try (XSSFWorkbook result = export(template,
                "{\"orders\":[{\"name\":\"钢笔\",\"quantity\":2,\"price\":12.5},"
                        + "{\"name\":\"纸\",\"quantity\":3,\"price\":8}]}")) {
            XSSFSheet sheet = result.getSheetAt(0);
            assertEquals(32, sheet.getRow(2).getHeightInPoints());
            assertEquals(6000, sheet.getColumnWidth(0));
            assertEquals(sheet.getRow(1).getCell(3).getCellStyle().getIndex(),
                    sheet.getRow(2).getCell(3).getCellStyle().getIndex());
            assertEquals(BorderStyle.THIN, sheet.getRow(2).getCell(3).getCellStyle().getBorderBottom());
            assertEquals("0.00", sheet.getRow(2).getCell(3).getCellStyle().getDataFormatString());
            assertEquals("C2*D2+$H$1", sheet.getRow(1).getCell(4).getCellFormula());
            assertEquals("C3*D3+$H$1", sheet.getRow(2).getCell(4).getCellFormula());
            assertEquals(3, sheet.getNumMergedRegions());
            assertTrue(sheet.getMergedRegions().stream().anyMatch(r -> r.formatAsString().equals("A3:B3")));
            assertTrue(sheet.getMergedRegions().stream().anyMatch(r -> r.formatAsString().equals("A4:B4")));
            assertEquals("尾部", text(sheet, 3, 0));
        }
    }

    @Test
    void removesEmptyListRowsAndTheirMergesWithoutLosingFollowingRows() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet sheet = book.createSheet("空数据");
            cells(sheet.createRow(0), "标题");
            cells(sheet.createRow(1), "${orders.name}");
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            cells(sheet.createRow(2), "${payments.amount}");
            cells(sheet.createRow(3), "结束：${name}");
        });
        try (XSSFWorkbook result = export(template,
                "{\"name\":\"张三\",\"orders\":[],\"payments\":[{\"amount\":10},{\"amount\":20}]}")) {
            XSSFSheet sheet = result.getSheetAt(0);
            assertEquals("标题", text(sheet, 0, 0));
            assertEquals(10, sheet.getRow(1).getCell(0).getNumericCellValue());
            assertEquals(20, sheet.getRow(2).getCell(0).getNumericCellValue());
            assertEquals("结束：张三", text(sheet, 3, 0));
            assertEquals(0, sheet.getNumMergedRegions());
            assertNull(sheet.getRow(4));
        }
    }

    @Test
    void handlesEmptyLastRowAndAllEmptyLists() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet sheet = book.createSheet("全空");
            cells(sheet.createRow(0), "${orders.name}");
            cells(sheet.createRow(1), "${tags}");
        });
        try (XSSFWorkbook result = export(template, "{\"orders\":[],\"tags\":[]}")) {
            assertEquals(0, result.getSheetAt(0).getPhysicalNumberOfRows());
        }
    }

    @Test
    void supportsNestedObjectPathsPrimitiveArraysAndMissingItemFields() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet sheet = book.createSheet("嵌套对象");
            cells(sheet.createRow(0), "${sales.orders.product.name}", "${sales.orders.quantity}");
            cells(sheet.createRow(1), "${tags}", "标签：${tags}");
        });
        try (XSSFWorkbook result = export(template,
                "{\"sales\":{\"orders\":[{\"product\":{\"name\":\"钢笔\"},\"quantity\":2},{}]},"
                        + "\"tags\":[\"加急\",\"送货\"]}")) {
            XSSFSheet sheet = result.getSheetAt(0);
            assertEquals("钢笔", text(sheet, 0, 0));
            assertEquals(CellType.BLANK, sheet.getRow(1).getCell(0).getCellType());
            assertEquals("加急", text(sheet, 2, 0));
            assertEquals("标签：送货", text(sheet, 3, 1));
        }
    }

    @Test
    void preservesLargeNumbersAndNumbersOutsideExcelRangeAsText() throws Exception {
        byte[] template = template(book -> cells(book.createSheet().createRow(0),
                "${id}", "${precise}", "${huge}", "${tiny}"));
        try (XSSFWorkbook result = export(template,
                "{\"id\":1234567890123456789,\"precise\":0.1234567890123456789,"
                        + "\"huge\":1e400,\"tiny\":1e-400}")) {
            XSSFRow row = result.getSheetAt(0).getRow(0);
            assertEquals("1234567890123456789", row.getCell(0).getStringCellValue());
            assertEquals("0.1234567890123456789", row.getCell(1).getStringCellValue());
            assertEquals(CellType.STRING, row.getCell(2).getCellType());
            assertEquals(CellType.STRING, row.getCell(3).getCellType());
        }
    }

    @Test
    void doesNotReinterpretDataAndEscapesRegexReplacement() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet sheet = book.createSheet();
            cells(sheet.createRow(0), "${orders.value}");
            cells(sheet.createRow(1), "值：${literal} / ${literal}", "${formula}");
        });
        try (XSSFWorkbook result = export(template,
                "{\"orders\":[{\"value\":\"${secret}\"},{\"value\":\"保留\"}],"
                        + "\"literal\":\"$5\\\\目录${secret}\",\"secret\":\"不要替换\",\"formula\":\"=1+1\"}")) {
            XSSFSheet sheet = result.getSheetAt(0);
            assertEquals("${secret}", text(sheet, 0, 0));
            assertEquals("值：$5\\目录${secret} / $5\\目录${secret}", text(sheet, 2, 0));
            assertEquals("=1+1", text(sheet, 2, 1));
            assertEquals(CellType.STRING, sheet.getRow(2).getCell(1).getCellType());
        }
    }

    @Test
    void rejectsDifferentArraysOnTheSameRowWithCellLocation() throws Exception {
        byte[] template = template(book -> cells(book.createSheet("订单").createRow(0),
                "${orders.name}", "${payments.amount}"));
        ExcelExportException error = assertThrows(ExcelExportException.class,
                () -> export(template, "{\"orders\":[],\"payments\":[]}"));
        assertTrue(error.getMessage().contains("B1"));
        assertTrue(error.getMessage().contains("orders 和 payments"));
    }

    @Test
    void rejectsNestedArraysAndWholeObjects() throws Exception {
        byte[] nested = template(book -> cells(book.createSheet().createRow(0), "${groups.members.name}"));
        assertThrows(ExcelExportException.class,
                () -> export(nested, "{\"groups\":[{\"members\":[{\"name\":\"张三\"}]}]}"));
        byte[] object = template(book -> cells(book.createSheet().createRow(0), "${customer}"));
        assertThrows(ExcelExportException.class, () -> export(object, "{\"customer\":{\"name\":\"张三\"}}"));
    }

    @Test
    void rejectsVerticalMergesOnListRows() throws Exception {
        byte[] template = template(book -> {
            XSSFSheet sheet = book.createSheet();
            cells(sheet.createRow(0), "${orders.name}");
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        });
        ExcelExportException error = assertThrows(ExcelExportException.class,
                () -> export(template, "{\"orders\":[{\"name\":\"钢笔\"}]}"));
        assertTrue(error.getMessage().contains("跨行合并"));
    }

    @Test
    void rejectsInvalidPathsAndWildcardPlaceholders() throws Exception {
        for (String placeholder : new String[]{"${}", "${customer..name}", "${orders.*}"}) {
            byte[] template = template(book -> cells(book.createSheet().createRow(0), placeholder));
            assertThrows(ExcelExportException.class, () -> export(template, "{}"));
        }
    }

    @Test
    void rejectsInvalidJsonAndNonObjectRootsBeforeWritingOutput() throws Exception {
        byte[] template = template(book -> book.createSheet());
        for (String json : new String[]{"[]", "null", "1", ""}) {
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            assertThrows(ExcelExportException.class,
                    () -> exporter.export(new ByteArrayInputStream(template), output, json));
            assertEquals(0, output.size());
        }
        assertThrows(JsonProcessingException.class, () -> export(template, "{broken}"));
        assertThrows(JsonProcessingException.class, () -> export(template, "{} {}"));
        assertThrows(JsonProcessingException.class, () -> export(template, "{\"name\":1,\"name\":2}"));
    }

    @Test
    void leavesCallerOwnedStreamsOpen() throws Exception {
        byte[] template = template(book -> cells(book.createSheet().createRow(0), "${name}"));
        TrackedInput input = new TrackedInput(template);
        TrackedOutput output = new TrackedOutput();
        exporter.export(input, output, "{\"name\":\"张三\"}");
        assertFalse(input.closed);
        assertFalse(output.closed);
        try (XSSFWorkbook result = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("张三", text(result.getSheetAt(0), 0, 0));
        }
    }

    @Test
    void reusesExporterWithoutLeakingPreviousData() throws Exception {
        byte[] template = template(book -> cells(book.createSheet().createRow(0), "${name}"));
        try (XSSFWorkbook first = export(template, "{\"name\":\"张三\"}");
             XSSFWorkbook second = export(template, "{\"name\":\"李四\"}")) {
            assertEquals("张三", text(first.getSheetAt(0), 0, 0));
            assertEquals("李四", text(second.getSheetAt(0), 0, 0));
        }
    }

    @Test
    void keepsTemplatesWithoutPlaceholders() throws Exception {
        byte[] template = template(book -> {
            cells(book.createSheet("原样保留").createRow(0), "静态内容");
            book.createSheet("空表");
        });
        try (XSSFWorkbook result = export(template, "{}")) {
            assertEquals("静态内容", text(result.getSheetAt(0), 0, 0));
            assertEquals(0, result.getSheetAt(1).getPhysicalNumberOfRows());
        }
    }

    private XSSFWorkbook export(byte[] template, String json) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        exporter.export(new ByteArrayInputStream(template), output, json);
        return new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()));
    }

    private byte[] template(TemplateBuilder builder) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            builder.build(workbook);
            workbook.write(output);
            return output.toByteArray();
        }
    }

    private static void cells(XSSFRow row, String... values) {
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    private String text(XSSFSheet sheet, int row, int column) {
        return sheet.getRow(row).getCell(column).getStringCellValue();
    }

    private interface TemplateBuilder {
        void build(XSSFWorkbook workbook);
    }

    private static final class TrackedInput extends ByteArrayInputStream {
        private boolean closed;

        private TrackedInput(byte[] data) {
            super(data);
        }

        @Override
        public void close() throws IOException {
            closed = true;
            super.close();
        }
    }

    private static final class TrackedOutput extends ByteArrayOutputStream {
        private boolean closed;

        @Override
        public void close() throws IOException {
            closed = true;
            super.close();
        }
    }
}
