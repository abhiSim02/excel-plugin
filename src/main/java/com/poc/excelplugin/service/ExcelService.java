package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.ExcelRequest;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;
import org.springframework.stereotype.Service;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.util.*;

@Slf4j
@Service
public class ExcelService {

    // Ideally, store this in application.properties or Vault. NEVER expose this.
    private static final String SECRET_KEY = "SUPER_SECRET_KEY_DO_NOT_SHARE";
    private static final String HMAC_ALGO = "HmacSHA256";

    public byte[] generateGenericExcel(ExcelRequest request) throws IOException {
        log.info("--- Service: Starting Excel Generation ---");

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Data");
            DataFormat poiDataFormat = workbook.createDataFormat();

            // 1. Generate a Unique ID for this specific file
            String fileId = UUID.randomUUID().toString();
            log.info("Generated File ID: {}", fileId);

            List<ExcelRequest.ColumnConfig> columns = request.getColumns();
            List<Map<String, Object>> dataList = request.getData();

            // 2. Create Headers
            CellStyle headerStyle = createHeaderStyle(workbook);
            Row headerRow = sheet.createRow(0);
            List<String> headerNames = new ArrayList<>();

            for (int i = 0; i < columns.size(); i++) {
                String header = columns.get(i).getHeader();
                headerNames.add(header);

                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header);
                cell.setCellStyle(headerStyle);
                int width = (columns.get(i).getWidth() != null) ? columns.get(i).getWidth() : 5000;
                sheet.setColumnWidth(i, width);
            }

            // 3. SIGN THE FILE (HMAC of File ID + Headers)
            // This binds the ID to the file structure.
            signWorkbook(workbook, headerNames, fileId);

            // 4. Pre-calculate Styles & Populate Data (Existing Logic)
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
                                try { cell.setCellValue(Double.parseDouble(value.toString())); }
                                catch (NumberFormatException e) { cell.setCellValue(value.toString()); }
                            } else {
                                cell.setCellValue(value.toString());
                            }
                        }
                    }
                }
            }

            // 5. Dropdowns
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
     * Calculates HMAC Signature based on File ID + Headers.
     */
    private void signWorkbook(XSSFWorkbook workbook, List<String> headerNames, String fileId) {
        try {
            // Data to sign: "UUID|ID|Car Name|City"
            String dataToSign = fileId + "|" + String.join("|", headerNames);
            String signature = calculateHMAC(dataToSign, SECRET_KEY);

            POIXMLProperties props = workbook.getProperties();
            POIXMLProperties.CustomProperties customProps = props.getCustomProperties();

            // Store ID and Signature
            if (customProps != null) {
                customProps.addProperty("X-File-ID", fileId);
                customProps.addProperty("X-Integrity-Sig", signature);
            }
            log.info("Signed Workbook. ID={}, Signature={}", fileId, signature);

        } catch (Exception e) {
            log.error("Failed to sign workbook", e);
            throw new RuntimeException("Could not sign file");
        }
    }

    /**
     * VERIFICATION LOGIC: Returns the File ID if valid, throws Exception if invalid.
     */
    public String verifyFileIntegrity(XSSFWorkbook workbook) {
        try {
            POIXMLProperties.CustomProperties props = workbook.getProperties().getCustomProperties();

            // 1. Extract the Stored Signature
            CTProperty sigProp = props.getProperty("X-Integrity-Sig");
            if (sigProp == null) throw new SecurityException("No signature found.");
            String storedSignature = sigProp.getLpwstr();

            // 2. Extract the File ID
            CTProperty idProp = props.getProperty("X-File-ID");
            if (idProp == null) throw new SecurityException("No File ID found.");
            String fileId = idProp.getLpwstr();

            // 3. Extract Headers
            XSSFSheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new SecurityException("No headers found.");
            List<String> extractedHeaders = new ArrayList<>();
            for (Cell cell : headerRow) extractedHeaders.add(cell.getStringCellValue());

            // 4. Re-calculate HMAC
            String dataToSign = fileId + "|" + String.join("|", extractedHeaders);
            String calculatedSignature = calculateHMAC(dataToSign, SECRET_KEY);

            log.info("Verifying ID: {}", fileId);

            if (!calculatedSignature.equals(storedSignature)) {
                log.warn("Signature Mismatch! File has been tampered with.");
                throw new SecurityException("Signature Mismatch");
            }

            return fileId; // Return the ID if valid

        } catch (SecurityException se) {
            throw se;
        } catch (Exception e) {
            log.error("Error verifying file", e);
            throw new RuntimeException("Error verifying file");
        }
    }

    private String calculateHMAC(String data, String key) throws NoSuchAlgorithmException, InvalidKeyException {
        SecretKeySpec secretKeySpec = new SecretKeySpec(key.getBytes(StandardCharsets.UTF_8), HMAC_ALGO);
        Mac mac = Mac.getInstance(HMAC_ALGO);
        mac.init(secretKeySpec);
        byte[] rawHmac = mac.doFinal(data.getBytes(StandardCharsets.UTF_8));
        return Base64.getEncoder().encodeToString(rawHmac);
    }

    // --- Helpers (Styles, Dropdowns, etc.) remain same ---
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