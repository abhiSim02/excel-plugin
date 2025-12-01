package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.ExcelRequest;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
@Service
public class ExcelService {

    // This is your secret key to validate files later
    private static final String PLATFORM_SECRET_KEY = "PLT_SECURE_SIG_V1";

    public byte[] generateGenericExcel(ExcelRequest request) throws IOException {
        log.info("--- Service: Starting Excel Generation ---");

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            // --- NEW: INJECT HIDDEN METADATA ---
            addPlatformSignature(workbook);

            XSSFSheet sheet = workbook.createSheet("Data");
            DataFormat poiDataFormat = workbook.createDataFormat();

            List<ExcelRequest.ColumnConfig> columns = request.getColumns();
            List<Map<String, Object>> dataList = request.getData();

            // 1. Pre-calculate Styles
            List<CellStyle> columnStyles = new ArrayList<>();
            for (ExcelRequest.ColumnConfig col : columns) {
                CellStyle style = workbook.createCellStyle();
                style.setLocked(!col.isEditable());
                if (col.isEditable()) {
                    style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                }
                if (col.getDataFormat() != null && !col.getDataFormat().isEmpty()) {
                    style.setDataFormat(poiDataFormat.getFormat(col.getDataFormat()));
                }
                columnStyles.add(style);
            }

            // 2. Create Headers
            CellStyle headerStyle = createHeaderStyle(workbook);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns.get(i).getHeader());
                cell.setCellStyle(headerStyle);
                int width = (columns.get(i).getWidth() != null) ? columns.get(i).getWidth() : 5000;
                sheet.setColumnWidth(i, width);
            }

            // 3. Populate Data
            int rowIdx = 1;
            if (dataList != null) {
                for (Map<String, Object> rowData : dataList) {
                    Row row = sheet.createRow(rowIdx++);
                    for (int colIdx = 0; colIdx < columns.size(); colIdx++) {
                        ExcelRequest.ColumnConfig colConfig = columns.get(colIdx);
                        Object value = rowData.get(colConfig.getKey());

                        Cell cell = row.createCell(colIdx);
                        cell.setCellStyle(columnStyles.get(colIdx));

                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value != null) {
                            if (colConfig.getDataFormat() != null && isNumeric(value.toString())) {
                                try {
                                    cell.setCellValue(Double.parseDouble(value.toString()));
                                } catch (NumberFormatException e) {
                                    cell.setCellValue(value.toString());
                                }
                            } else {
                                cell.setCellValue(value.toString());
                            }
                        }
                    }
                }
            }

            // 4. Dropdowns
            for (int colIdx = 0; colIdx < columns.size(); colIdx++) {
                ExcelRequest.ColumnConfig col = columns.get(colIdx);
                if (col.getDropdown() != null && !col.getDropdown().isEmpty()) {
                    String[] options = col.getDropdown().toArray(new String[0]);
                    addDropdownValidation(sheet, options, colIdx, 1, Math.max(100, rowIdx + 100));
                }
            }

            sheet.protectSheet("password123");

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    /**
     * Adds hidden custom properties to the Excel file.
     * We add it in two places to ensure it's verifiable.
     */
    private void addPlatformSignature(XSSFWorkbook workbook) {
        POIXMLProperties props = workbook.getProperties();

        // 1. Custom Property (The original hidden location)
        // Access via File > Info > Properties > Advanced Properties > Custom
        POIXMLProperties.CustomProperties customProps = props.getCustomProperties();
        if (customProps != null) {
            customProps.addProperty("X-Platform-Auth", PLATFORM_SECRET_KEY);
        }

        // 2. Core Property - Category (Backup location)
        // This is often visible in the standard "Tags" or "Category" field in file properties
        // Access via File > Info > Show All Properties
        POIXMLProperties.CoreProperties coreProps = props.getCoreProperties();
        if (coreProps != null) {
            coreProps.setCategory("Auth: " + PLATFORM_SECRET_KEY);
            coreProps.setCreator("Internal Excel Plugin");
        }

        log.info("Injected Metadata Signature (Custom & Category) into Workbook");
    }

    // --- Helpers ---

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setLocked(true);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private void addDropdownValidation(XSSFSheet sheet, String[] options, int colIndex, int firstRow, int lastRow) {
        try {
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
            DataValidationConstraint constraint = dvHelper.createExplicitListConstraint(options);
            CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, colIndex, colIndex);
            DataValidation validation = dvHelper.createValidation(constraint, addressList);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);
        } catch (Exception e) {
            log.error("Error adding dropdown", e);
        }
    }

    private boolean isNumeric(String str) {
        if (str == null) return false;
        try { Double.parseDouble(str); return true; } catch (NumberFormatException e) { return false; }
    }
}