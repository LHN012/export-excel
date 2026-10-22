package com.example.excel;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** 列表行插入/删除及复制。由调用者按原始行号从下向上处理。 */
final class ListRenderer {
    private final ScalarRenderer scalarRenderer;

    ListRenderer(ScalarRenderer scalarRenderer) {
        this.scalarRenderer = scalarRenderer;
    }

    void validate(XSSFSheet sheet, int rowIndex) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.getFirstRow() <= rowIndex && region.getLastRow() >= rowIndex
                    && region.getFirstRow() != region.getLastRow()) {
                throw new ExcelExportException("工作表「" + sheet.getSheetName() + "」第 "
                        + (rowIndex + 1) + " 行是列表模板行，不能跨行合并单元格："
                        + region.formatAsString());
            }
        }
    }

    void render(XSSFSheet sheet, int rowIndex, JsonNode root, String listPath, JsonNode list) {
        int count = list.size();
        XSSFRow templateRow = sheet.getRow(rowIndex);
        if (count == 0) {
            removeRow(sheet, templateRow);
            return;
        }
        int extra = count - 1;
        if ((long) sheet.getLastRowNum() + extra > SpreadsheetVersion.EXCEL2007.getLastRowIndex()) {
            throw new ExcelExportException("工作表「" + sheet.getSheetName()
                    + "」展开 " + listPath + " 后超过 XLSX 的最大行数");
        }
        if (extra > 0 && rowIndex < sheet.getLastRowNum()) {
            sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), extra, true, false);
        }

        // 先完成所有复制，再填值，避免第一条数据覆盖占位符后被继续复制。
        CellCopyPolicy policy = new CellCopyPolicy();
        policy.setCopyMergedRegions(true);
        for (int i = 1; i < count; i++) {
            sheet.copyRows(rowIndex, rowIndex, rowIndex + i, policy);
            XSSFRow copy = sheet.getRow(rowIndex + i);
            copy.setZeroHeight(templateRow.getZeroHeight());
            if (templateRow.isFormatted()) {
                copy.setRowStyle(templateRow.getRowStyle());
            }
        }
        for (int i = 0; i < count; i++) {
            scalarRenderer.render(sheet.getRow(rowIndex + i), root, listPath, list.get(i));
        }
    }

    private void removeRow(XSSFSheet sheet, XSSFRow row) {
        int rowIndex = row.getRowNum();
        // POI 删除行时不会自动删除这一行自身的合并区域。
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() == rowIndex && region.getLastRow() == rowIndex) {
                sheet.removeMergedRegion(i);
            }
        }
        int lastRow = sheet.getLastRowNum();
        sheet.removeRow(row);
        if (rowIndex < lastRow) {
            sheet.shiftRows(rowIndex + 1, lastRow, -1, true, false);
        }
    }
}
