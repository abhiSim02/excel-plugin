package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.ExcelRequest;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;
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
    private static final String REGEX_DELIMITER = "\\|";

    // --- Wrapper to return both File Bytes AND the Signature to the Controller ---
    @Data
    @AllArgsConstructor
    public static class GenerationResult {
        private byte[] fileContent;
        private String signature; // The 24-char hash to save in DB
    }

    public GenerationResult generateGenericExcel(ExcelRequest request) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            String rawEntity = (request.getEntityName() != null) ? request.getEntityName() : "Data";
            String safeSheetName = rawEntity.replaceAll("[^a-zA-Z0-9 ]", "_");
            if (safeSheetName.length() > 31) safeSheetName = safeSheetName.substring(0, 31);

            XSSFSheet sheet = workbook.createSheet(safeSheetName);
            DataFormat poiDataFormat = workbook.createDataFormat();

            // 1. Generate Metadata & Signature
            String entity = request.getEntityName() != null ? request.getEntityName() : "unknown";
            String user = request.getUserId() != null ? request.getUserId() : "unknown";
            // Use readable timestamp
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"));

            // This helper calculates the signature, injects metadata, and RETURNS the signature
            String generatedSignature = signWorkbookAndGetSignature(workbook, entity, user, timestamp);

            // 2. Populate Columns & Data
            List<ExcelRequest.ColumnConfig> columns = request.getColumns();
            List<Map<String, Object>> dataList = request.getData();

            CellStyle headerStyle = createHeaderStyle(workbook);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns.get(i).getHeader());
                cell.setCellStyle(headerStyle);
                int width = (columns.get(i).getWidth() != null) ? columns.get(i).getWidth() : 5000;
                sheet.setColumnWidth(i, width);
            }

            // Styles & Data Rows
            List<CellStyle> columnStyles = new ArrayList<>();
            for (ExcelRequest.ColumnConfig col : columns) {
                CellStyle style = workbook.createCellStyle();
                style.setLocked(!col.isEditable());
                if (col.isEditable()) {
                    style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
                if (col.getDataFormat() != null && !col.getDataFormat().isEmpty()) {
                    style.setDataFormat(poiDataFormat.getFormat(col.getDataFormat()));
                }
                columnStyles.add(style);
            }

            int rowIdx = 1;
            if (dataList != null) {
                for (Map<String, Object> rowData : dataList) {
                    Row row = sheet.createRow(rowIdx++);
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

            // Dropdowns
            for (int colIdx = 0; colIdx < columns.size(); colIdx++) {
                ExcelRequest.ColumnConfig col = columns.get(colIdx);
                if (col.getDropdown() != null && !col.getDropdown().isEmpty()) {
                    String[] options = col.getDropdown().toArray(new String[0]);
                    addDropdownValidation(sheet, options, colIdx, 1, Math.max(100, rowIdx + 100));
                }
            }

            sheet.protectSheet(sheetPassword);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);

            // Return BOTH content (for file) and signature (for DB)
            return new GenerationResult(out.toByteArray(), generatedSignature);
        }
    }

    private String signWorkbookAndGetSignature(XSSFWorkbook workbook, String entity, String user, String timestamp) {
        try {
            // Data to sign: "Entity|User|Timestamp"
            String dataPayload = entity + DELIMITER + user + DELIMITER + timestamp;
            String fullSignature = calculateHMAC(dataPayload, secretKey);

            // Truncate to 24 chars
            String signature24 = fullSignature.length() > 24 ? fullSignature.substring(0, 24) : fullSignature;

            // Combine: "Entity|User|Timestamp|Signature"
            String platformKey = dataPayload + DELIMITER + signature24;

            POIXMLProperties props = workbook.getProperties();
            POIXMLProperties.CustomProperties customProps = props.getCustomProperties();
            if (customProps != null) {
                customProps.addProperty("platform_key", platformKey);
            }
            log.info("Generated Platform Key: {}", platformKey);

            return signature24; // Return only the hash to be stored in DB

        } catch (Exception e) {
            throw new RuntimeException("Could not sign file", e);
        }
    }

    /**
     * Extracts the full key from the file for the Controller to process.
     */
    public String extractPlatformKey(XSSFWorkbook workbook) {
        try {
            POIXMLProperties.CustomProperties props = workbook.getProperties().getCustomProperties();
            CTProperty tokenProp = props.getProperty("platform_key");
            if (tokenProp == null) throw new SecurityException("Missing Hash Key (platform_key)");
            return tokenProp.getLpwstr();
        } catch (Exception e) {
            throw new SecurityException("Could not read platform key: " + e.getMessage());
        }
    }

    // Double check the math (Internal Integrity check)
    public void verifyMathIntegrity(String entity, String user, String timestamp, String storedSig) {
        try {
            String dataPayload = entity + DELIMITER + user + DELIMITER + timestamp;
            String fullSignature = calculateHMAC(dataPayload, secretKey);
            String calculatedSig24 = fullSignature.length() > 24 ? fullSignature.substring(0, 24) : fullSignature;

            if (!calculatedSig24.equals(storedSig)) {
                throw new SecurityException("Signature Mismatch - Metadata modified");
            }
        } catch (Exception e) {
            throw new RuntimeException("Math verification failed", e);
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
}