package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.ExcelRequest;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Slf4j
@Service
public class ExcelService {

    @Value("${excel.security.secret}")
    private String secretKey;

    @Value("${excel.security.algorithm}")
    private String hmacAlgo;

    @Value("${excel.sheet.password}")
    private String sheetPassword;

    private static final String DELIMITER = "|";
    // Sheets
    private static final String HIDDEN_SHEET_METADATA = "metadata_protected";
    private static final String HIDDEN_SHEET_LOOKUP = "lookup_data";

    @Data
    @AllArgsConstructor
    public static class GenerationResult {
        private byte[] fileContent;
        private String signature;
    }

    public GenerationResult generateGenericExcel(ExcelRequest request) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            // 1. Create Main Data Sheet
            String rawEntity = (request.getEntityName() != null) ? request.getEntityName() : "Data";
            String safeSheetName = rawEntity.replaceAll("[^a-zA-Z0-9 ]", "_");
            if (safeSheetName.length() > 31) safeSheetName = safeSheetName.substring(0, 31);

            XSSFSheet mainSheet = workbook.createSheet(safeSheetName);
            DataFormat poiDataFormat = workbook.createDataFormat();

            // 2. Security: Sign File
            String entity = request.getEntityName() != null ? request.getEntityName() : "unknown";
            String user = request.getUserId() != null ? request.getUserId() : "unknown";
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"));
            String generatedSignature = signWorkbookWithHiddenSheet(workbook, entity, user, timestamp);

            // 3. Prepare Lookup Sheet for Dropdowns
            XSSFSheet lookupSheet = workbook.createSheet(HIDDEN_SHEET_LOOKUP);

            // 4. Headers
            List<ExcelRequest.ColumnConfig> columns = request.getColumns();
            CellStyle headerStyle = createHeaderStyle(workbook);
            Row headerRow = mainSheet.createRow(0);
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns.get(i).getHeader());
                cell.setCellStyle(headerStyle);
                int width = (columns.get(i).getWidth() != null) ? columns.get(i).getWidth() : 5000;
                mainSheet.setColumnWidth(i, width);
            }

            // 5. Styles & Configs
            List<CellStyle> columnStyles = new ArrayList<>();
            int lookupColIdx = 0; // Track column index in the hidden lookup sheet

            for (int i = 0; i < columns.size(); i++) {
                ExcelRequest.ColumnConfig colConfig = columns.get(i);

                // Style Setup
                CellStyle style = workbook.createCellStyle();
                style.setLocked(!colConfig.isEditable());
                if (colConfig.isEditable()) {
                    style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
                if (colConfig.getDataFormat() != null && !colConfig.getDataFormat().isEmpty()) {
                    style.setDataFormat(poiDataFormat.getFormat(colConfig.getDataFormat()));
                }
                columnStyles.add(style);

                // --- ROBUST DROPDOWN & RED HIGHLIGHT LOGIC ---
                if (colConfig.getDropdown() != null && !colConfig.getDropdown().isEmpty()) {
                    // A. Write Options to Hidden Sheet
                    String namedRangeName = "List_" + colConfig.getKey();
                    // Clean name (Named ranges can't have spaces)
                    namedRangeName = namedRangeName.replaceAll("[^a-zA-Z0-9_]", "_");

                    writeLookupData(lookupSheet, colConfig.getDropdown(), lookupColIdx);

                    // B. Create Named Range (e.g., "List_Region")
                    createNamedRange(workbook, HIDDEN_SHEET_LOOKUP, namedRangeName, lookupColIdx, colConfig.getDropdown().size());

                    // C. Apply Data Validation (Using Name)
                    applyDataValidation(mainSheet, namedRangeName, i);

                    // D. Apply Conditional Formatting (Red if invalid)
                    applyErrorHighlighting(mainSheet, i, namedRangeName);

                    lookupColIdx++;
                }
            }

            // 6. Populate Data
            List<Map<String, Object>> dataList = request.getData();
            int rowIdx = 1;
            if (dataList != null) {
                for (Map<String, Object> rowData : dataList) {
                    Row row = mainSheet.createRow(rowIdx++);
                    for (int colIdx = 0; colIdx < columns.size(); colIdx++) {
                        Object value = rowData.get(columns.get(colIdx).getKey());
                        Cell cell = row.createCell(colIdx);
                        cell.setCellStyle(columnStyles.get(colIdx));

                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value != null) {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
            }

            // 7. Finalize (Hide Sheets & Protect)
            workbook.setSheetVisibility(workbook.getSheetIndex(lookupSheet), SheetVisibility.VERY_HIDDEN);
            mainSheet.protectSheet(sheetPassword);
            workbook.setActiveSheet(workbook.getSheetIndex(mainSheet));

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);

            return new GenerationResult(out.toByteArray(), generatedSignature);
        }
    }

    // --- NEW: Write Dropdown Data to Hidden Sheet ---
    private void writeLookupData(XSSFSheet lookupSheet, List<String> options, int colIdx) {
        for (int i = 0; i < options.size(); i++) {
            Row row = lookupSheet.getRow(i);
            if (row == null) row = lookupSheet.createRow(i);
            Cell cell = row.createCell(colIdx);
            cell.setCellValue(options.get(i));
        }
    }

    // --- NEW: Create Excel Named Range ---
    private void createNamedRange(Workbook workbook, String sheetName, String rangeName, int colIdx, int rowCount) {
        Name namedRange = workbook.createName();
        namedRange.setNameName(rangeName);
        String colLetter = CellReference.convertNumToColString(colIdx);
        // Formula: lookup_data!$A$1:$A$5
        String reference = sheetName + "!$" + colLetter + "$1:$" + colLetter + "$" + rowCount;
        namedRange.setRefersToFormula(reference);
    }

    // --- NEW: Apply Data Validation using Named Range ---
    private void applyDataValidation(XSSFSheet sheet, String namedRangeName, int colIdx) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        // Constraint: =List_Region
        DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(namedRangeName);
        // Apply to rows 1 to 5000
        CellRangeAddressList addressList = new CellRangeAddressList(1, 5000, colIdx, colIdx);
        DataValidation validation = dvHelper.createValidation(constraint, addressList);
        validation.setShowErrorBox(true);
        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.createErrorBox("Invalid Input", "Please select a valid value from the dropdown list.");
        sheet.addValidationData(validation);
    }

    // --- NEW: Apply RED Background if Value NOT in List ---
    private void applyErrorHighlighting(XSSFSheet sheet, int colIdx, String namedRangeName) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Formula: =AND(A2<>"", COUNTIF(List_Region, A2)=0)
        // If cell is not empty AND count in list is 0 -> Error
        String colLetter = CellReference.convertNumToColString(colIdx);
        String ruleFormula = String.format("AND(%s2<>\"\", COUNTIF(%s, %s2)=0)", colLetter, namedRangeName, colLetter);

        ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(ruleFormula);
        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(IndexedColors.RED.getIndex());
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // Font Color White (Optional, for readability on Red)
        FontFormatting font = rule.createFontFormatting();
        font.setFontColorIndex(IndexedColors.WHITE.getIndex());

        // Apply to range (Row 2 to 5000)
        CellRangeAddress[] regions = {
                new CellRangeAddress(1, 5000, colIdx, colIdx)
        };
        sheetCF.addConditionalFormatting(regions, rule);
    }

    // --- Existing Security Logic (Unchanged) ---
    private String signWorkbookWithHiddenSheet(XSSFWorkbook workbook, String entity, String user, String timestamp) {
        try {
            String dataPayload = entity + DELIMITER + user + DELIMITER + timestamp;
            String fullSignature = calculateHMAC(dataPayload, secretKey);
            String signature24 = fullSignature.length() > 24 ? fullSignature.substring(0, 24) : fullSignature;
            String platformKey = dataPayload + DELIMITER + signature24;
            String encodedKey = Base64.getEncoder().encodeToString(platformKey.getBytes(StandardCharsets.UTF_8));

            XSSFSheet metaSheet = workbook.createSheet(HIDDEN_SHEET_METADATA);
            Row row = metaSheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(encodedKey);
            metaSheet.protectSheet(sheetPassword);
            workbook.setSheetVisibility(workbook.getSheetIndex(metaSheet), SheetVisibility.VERY_HIDDEN);

            return signature24;
        } catch (Exception e) {
            throw new RuntimeException("Could not sign file", e);
        }
    }

    public String extractPlatformKey(XSSFWorkbook workbook) {
        try {
            XSSFSheet metaSheet = workbook.getSheet(HIDDEN_SHEET_METADATA);
            if (metaSheet == null) throw new SecurityException("Missing Metadata Sheet");
            String encodedKey = metaSheet.getRow(0).getCell(0).getStringCellValue();
            return new String(Base64.getDecoder().decode(encodedKey), StandardCharsets.UTF_8);
        } catch (Exception e) {
            throw new SecurityException("Could not read platform key: " + e.getMessage());
        }
    }

    private String calculateHMAC(String data, String key) throws NoSuchAlgorithmException, InvalidKeyException {
        SecretKeySpec secretKeySpec = new SecretKeySpec(key.getBytes(StandardCharsets.UTF_8), hmacAlgo);
        Mac mac = Mac.getInstance(hmacAlgo);
        mac.init(secretKeySpec);
        return Base64.getEncoder().encodeToString(mac.doFinal(data.getBytes(StandardCharsets.UTF_8)));
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}