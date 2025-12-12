package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.ExcelRequest;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.codec.digest.DigestUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.*;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;

@Slf4j
@Service
public class ExcelService {

    @Value("${excel.security.secret}")
    private String secretKey;

    @Value("${excel.sheet.password}")
    private String sheetPassword;

    private static final String HIDDEN_SHEET_METADATA = "metadata_protected";
    private static final String HIDDEN_SHEET_LOOKUP = "lookup_data";
    private static final String HIDDEN_SHEET_SHADOW = "shadow_data";
    private static final int MAX_ROWS = SpreadsheetVersion.EXCEL2007.getLastRowIndex();

    @Data
    @AllArgsConstructor
    public static class GenerationResult {
        private byte[] fileContent;
        private String signature;
    }

    public GenerationResult generateLargeExcel(ExcelRequest request, String timestamp) throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            workbook.setCompressTempFiles(true);

            // 1. Create Sheets
            SXSSFSheet mainSheet = workbook.createSheet(sanitizeSheetName(request.getEntityName()));
            SXSSFSheet shadowSheet = workbook.createSheet(HIDDEN_SHEET_SHADOW);
            SXSSFSheet lookupSheet = workbook.createSheet(HIDDEN_SHEET_LOOKUP);

            mainSheet.trackAllColumnsForAutoSizing();

            // 2. Setup Styles & Headers
            Map<String, CellStyle> styles = createStyles(workbook);
            createHeaderRow(mainSheet, request.getColumns(), styles.get("header"));
            createHeaderRow(shadowSheet, request.getColumns(), styles.get("header"));

            // --- FIX START: Apply Column-Level Styling ---
            // This ensures empty rows below the data inherit the correct Lock/Unlock status
            for (int i = 0; i < request.getColumns().size(); i++) {
                ExcelRequest.ColumnConfig col = request.getColumns().get(i);
                if (col.isEditable()) {
                    // If editable, the WHOLE column is unlocked by default (even empty rows)
                    mainSheet.setDefaultColumnStyle(i, styles.get("editable"));
                } else {
                    // Otherwise, the whole column is locked
                    mainSheet.setDefaultColumnStyle(i, styles.get("locked"));
                }
            }
            // --- FIX END ---

            // 3. Configure Logic
            configureSheetLogic(mainSheet, workbook, request.getColumns(), lookupSheet);

            // 4. Stream Data & Calculate Content Hash
            StringBuilder contentHashBuilder = new StringBuilder();

            if (request.getData() != null) {
                int rowIdx = 1;
                for (Map<String, Object> rowData : request.getData()) {
                    Row mainRow = mainSheet.createRow(rowIdx);
                    Row shadowRow = shadowSheet.createRow(rowIdx);

                    for (int colIdx = 0; colIdx < request.getColumns().size(); colIdx++) {
                        ExcelRequest.ColumnConfig col = request.getColumns().get(colIdx);
                        Object value = rowData.get(col.getKey());

                        // A. Write to Visible Main Sheet
                        Cell mainCell = mainRow.createCell(colIdx);
                        setCellValue(mainCell, value);
                        // Note: setCellStyle is strictly not required here anymore because
                        // we set the default column style above, but it's good practice to keep it
                        // to ensure specific formatting (dates/numbers) isn't lost.
                        mainCell.setCellStyle(col.isEditable() ? styles.get("editable") : styles.get("locked"));

                        // B. Write to Hidden Shadow Sheet
                        Cell shadowCell = shadowRow.createCell(colIdx);
                        setCellValue(shadowCell, value);

                        // C. Hashing Logic (Read-Only Columns only)
                        if (!col.isEditable() && value != null) {
                            contentHashBuilder.append(value.toString().trim());
                        }
                    }
                    rowIdx++;
                    if (rowIdx >= MAX_ROWS) break;
                }
            }

            // 5. Finalize Hash & Sign
            String dataHash = DigestUtils.sha256Hex(contentHashBuilder.toString());
            String signature = signWorkbook(workbook, request.getEntityName(), request.getUserId(), timestamp, dataHash);

            // 6. Hide Sheets & Protect
            workbook.setSheetVisibility(workbook.getSheetIndex(lookupSheet), SheetVisibility.VERY_HIDDEN);
            workbook.setSheetVisibility(workbook.getSheetIndex(shadowSheet), SheetVisibility.VERY_HIDDEN);

            // Protect the sheet - now editable columns will remain editable in new rows
            mainSheet.protectSheet(sheetPassword);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            workbook.dispose();

            return new GenerationResult(out.toByteArray(), signature);
        }
    }

    private void configureSheetLogic(SXSSFSheet sheet, SXSSFWorkbook workbook, List<ExcelRequest.ColumnConfig> columns, Sheet lookupSheet) {
        XSSFSheet xssfSheet = workbook.getXSSFWorkbook().getSheet(sheet.getSheetName());
        SheetConditionalFormatting sheetCF = xssfSheet.getSheetConditionalFormatting();

        int lookupColIdx = 0;

        for (int i = 0; i < columns.size(); i++) {
            ExcelRequest.ColumnConfig col = columns.get(i);
            String colLetter = CellReference.convertNumToColString(i);
            CellRangeAddress[] fullColumnRegion = { new CellRangeAddress(1, MAX_ROWS - 1, i, i) };

            // --- 1. RED Logic (Validation Error) ---
            if (col.getDropdown() != null && !col.getDropdown().isEmpty()) {
                writeLookupData(lookupSheet, col.getDropdown(), lookupColIdx);

                String rangeName = "List_" + col.getKey().replaceAll("[^a-zA-Z0-9_]", "_");
                Name namedRange = workbook.createName();
                namedRange.setNameName(rangeName);
                String lookupColLetter = CellReference.convertNumToColString(lookupColIdx);
                namedRange.setRefersToFormula(HIDDEN_SHEET_LOOKUP + "!$" + lookupColLetter + "$1:$" + lookupColLetter + "$" + col.getDropdown().size());

                DataValidationHelper dvHelper = sheet.getDataValidationHelper();
                DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(rangeName);
                DataValidation validation = dvHelper.createValidation(constraint, new CellRangeAddressList(1, MAX_ROWS - 1, i, i));
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);

                String redFormula = String.format("AND(%s2<>\"\", COUNTIF(%s, %s2)=0)", colLetter, rangeName, colLetter);
                ConditionalFormattingRule redRule = sheetCF.createConditionalFormattingRule(redFormula);
                PatternFormatting redFill = redRule.createPatternFormatting();
                redFill.setFillBackgroundColor(IndexedColors.RED.getIndex());
                redFill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
                FontFormatting whiteFont = redRule.createFontFormatting();
                whiteFont.setFontColorIndex(IndexedColors.WHITE.getIndex());

                sheetCF.addConditionalFormatting(fullColumnRegion, redRule);
                lookupColIdx++;
            }

            // --- 2. ORANGE Logic (Modified) ---
            if (col.isEditable()) {
                String orangeFormula = String.format("%s2<>%s!%s2", colLetter, HIDDEN_SHEET_SHADOW, colLetter);
                ConditionalFormattingRule orangeRule = sheetCF.createConditionalFormattingRule(orangeFormula);
                PatternFormatting orangeFill = orangeRule.createPatternFormatting();
                orangeFill.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());
                orangeFill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

                sheetCF.addConditionalFormatting(fullColumnRegion, orangeRule);
            }
        }
    }

    private String signWorkbook(Workbook workbook, String entity, String user, String timestamp, String dataHash) {
        try {
            String payload = entity + "|" + user + "|" + timestamp + "|" + dataHash;
            Mac mac = Mac.getInstance("HmacSHA256");
            SecretKeySpec secretKeySpec = new SecretKeySpec(secretKey.getBytes(StandardCharsets.UTF_8), "HmacSHA256");
            mac.init(secretKeySpec);

            String fullSignature = Base64.getEncoder().encodeToString(mac.doFinal(payload.getBytes(StandardCharsets.UTF_8)));

            // --- MODIFIED: TRUNCATE TO 24 CHARS ---
            String signature = fullSignature.length() > 24 ? fullSignature.substring(0, 24) : fullSignature;

            // Store in Hidden Metadata Sheet
            String storedValue = signature + "::" + dataHash + "::" + entity + "::" + user + "::" + timestamp;
            Sheet metaSheet = workbook.createSheet(HIDDEN_SHEET_METADATA);
            metaSheet.createRow(0).createCell(0).setCellValue(Base64.getEncoder().encodeToString(storedValue.getBytes(StandardCharsets.UTF_8)));

            workbook.setSheetVisibility(workbook.getSheetIndex(metaSheet), SheetVisibility.VERY_HIDDEN);
            return signature;
        } catch (Exception e) {
            throw new RuntimeException("Signing failed", e);
        }
    }

    private void writeLookupData(Sheet lookupSheet, List<String> options, int colIdx) {
        for (int i = 0; i < options.size(); i++) {
            Row row = lookupSheet.getRow(i);
            if (row == null) row = lookupSheet.createRow(i);
            row.createCell(colIdx).setCellValue(options.get(i));
        }
    }

    private void createHeaderRow(Sheet sheet, List<ExcelRequest.ColumnConfig> columns, CellStyle style) {
        Row row = sheet.createRow(0);
        for (int i = 0; i < columns.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columns.get(i).getHeader());
            cell.setCellStyle(style);
            sheet.setColumnWidth(i, columns.get(i).getWidth() != null ? columns.get(i).getWidth() : 4000);
        }
    }

    private void setCellValue(Cell cell, Object value) {
        if (value instanceof Number) cell.setCellValue(((Number) value).doubleValue());
        else if (value != null) cell.setCellValue(value.toString());
    }

    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();

        CellStyle header = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        header.setFont(headerFont);
        header.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("header", header);

        CellStyle editable = wb.createCellStyle();
        editable.setLocked(false); // <--- CRITICAL: Must be false
        editable.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        editable.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("editable", editable);

        CellStyle locked = wb.createCellStyle();
        locked.setLocked(true); // Default is true, but explicit is safer
        styles.put("locked", locked);

        return styles;
    }

    private String sanitizeSheetName(String name) {
        return name == null ? "Data" : name.replaceAll("[^a-zA-Z0-9 ]", "_");
    }

    public Map<String, Object> analyzeChanges(MultipartFile file) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream())) {

            Sheet mainSheet = workbook.getSheetAt(0);
            Sheet shadowSheet = workbook.getSheet("shadow_data");

            if (shadowSheet == null) {
                throw new IllegalArgumentException("Invalid File: Missing shadow_data for audit.");
            }

            // 1. Pre-load Validation Rules
            Map<Integer, Set<String>> validationRules = extractValidationRules(workbook, mainSheet);

            int totalChangedCells = 0;
            int totalInvalidCells = 0;

            // Track unique row indices for different stats
            Set<Integer> rowsWithChanges = new HashSet<>();
            Set<Integer> rowsWithErrors = new HashSet<>();

            List<Map<String, Object>> rowDetails = new ArrayList<>();

            // 2. Iterate Rows
            for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
                Row mainRow = mainSheet.getRow(i);
                Row shadowRow = shadowSheet.getRow(i);

                if (mainRow == null) continue;

                boolean rowChanged = false;
                boolean rowInvalid = false;

                Map<String, Object> details = new LinkedHashMap<>();
                details.put("rowIndex", i + 1);

                // 3. Iterate Columns
                for (int j = 0; j < mainRow.getLastCellNum(); j++) {
                    String original = getCellValueAsString(shadowRow.getCell(j));
                    String current = getCellValueAsString(mainRow.getCell(j));

                    // A. CHANGE DETECTION
                    if (!Objects.equals(original, current)) {
                        details.put("col_" + j + "_change", "Changed from [" + original + "] to [" + current + "]");
                        totalChangedCells++;
                        rowChanged = true;
                    }

                    // B. VALIDATION CHECK
                    if (validationRules.containsKey(j)) {
                        Set<String> allowedValues = validationRules.get(j);
                        if (!current.isEmpty() && !allowedValues.contains(current)) {
                            details.put("col_" + j + "_error", "Invalid Value: [" + current + "]");
                            totalInvalidCells++;
                            rowInvalid = true;
                        }
                    }
                }

                if (rowChanged) rowsWithChanges.add(i);
                if (rowInvalid) rowsWithErrors.add(i);

                // Report if EITHER changed OR invalid
                if (rowChanged || rowInvalid) {
                    // Add status flag for UI convenience
                    details.put("status", rowInvalid ? "INVALID" : "MODIFIED_VALID");
                    rowDetails.add(details);
                }
            }

            // 4. Construct Precise Stats
            int totalRows = mainSheet.getLastRowNum();

            Map<String, Object> stats = new LinkedHashMap<>();
            stats.put("totalRowsProcessed", totalRows);

            // Purely valid data (Total - Invalid)
            stats.put("totalValidRows", totalRows - rowsWithErrors.size());

            // The number you are looking for (3)
            stats.put("totalRowsWithErrors", rowsWithErrors.size());

            // The number of touched rows (6)
            stats.put("totalRowsModified", rowsWithChanges.size());

            stats.put("totalChangedCells", totalChangedCells);
            stats.put("totalInvalidCells", totalInvalidCells);

            if (totalRows > 0) {
                stats.put("errorRate", String.format("%.2f%%", (double) rowsWithErrors.size() / totalRows * 100));
            } else {
                stats.put("errorRate", "0.00%");
            }

            Map<String, Object> result = new LinkedHashMap<>();
            result.put("statistics", stats);
            result.put("details", rowDetails);

            return result;
        }
    }
    /**
     * Extracts dropdown options from the file by resolving Named Ranges.
     */
    private Map<Integer, Set<String>> extractValidationRules(Workbook workbook, Sheet mainSheet) {
        Map<Integer, Set<String>> rules = new HashMap<>();

        // Get all data validations from the main sheet
        List<? extends DataValidation> validations = mainSheet.getDataValidations();

        for (DataValidation dv : validations) {
            // Get the Named Range name (e.g., "List_Region")
            String formula = dv.getValidationConstraint().getFormula1();

            if (formula != null && !formula.isEmpty()) {
                // Resolve Name to Range (e.g., lookup_data!$A$1:$A$5)
                Name namedRange = workbook.getName(formula);
                if (namedRange != null) {
                    Set<String> allowedValues = getValuesFromNamedRange(workbook, namedRange);

                    // Map this rule to the columns it applies to
                    CellRangeAddress[] regions = dv.getRegions().getCellRangeAddresses();
                    for (CellRangeAddress region : regions) {
                        for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
                            rules.put(col, allowedValues);
                        }
                    }
                }
            }
        }
        return rules;
    }

    private Set<String> getValuesFromNamedRange(Workbook workbook, Name namedRange) {
        Set<String> values = new HashSet<>();
        try {
            // Get the reference (e.g., lookup_data!$A$1:$A$5)
            String reference = namedRange.getRefersToFormula();
            AreaReference area = new AreaReference(reference, SpreadsheetVersion.EXCEL2007);

            // Parse Sheet Name and Cells
            CellReference[] cells = area.getAllReferencedCells();
            if (cells.length > 0) {
                String sheetName = cells[0].getSheetName();
                Sheet lookupSheet = workbook.getSheet(sheetName);

                for (CellReference cellRef : cells) {
                    Row row = lookupSheet.getRow(cellRef.getRow());
                    if (row != null) {
                        Cell cell = row.getCell(cellRef.getCol());
                        String val = getCellValueAsString(cell);
                        if (!val.isEmpty()) {
                            values.add(val); // Strings are case-sensitive usually
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.warn("Could not resolve named range: " + namedRange.getNameName(), e);
        }
        return values;
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }
}