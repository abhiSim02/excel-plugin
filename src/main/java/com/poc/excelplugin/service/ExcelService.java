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
import org.springframework.beans.factory.annotation.Value;
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

    // Injecting values from application.properties
    @Value("${excel.security.secret}")
    private String secretKey;

    @Value("${excel.security.algorithm}")
    private String hmacAlgo;

    @Value("${excel.sheet.password}")
    private String sheetPassword;

    public byte[] generateGenericExcel(ExcelRequest request) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Data");
            DataFormat poiDataFormat = workbook.createDataFormat();

            // 1. Generate Headers & Data
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

            // 2. SIGN THE FILE
            String fileId = UUID.randomUUID().toString();
            signWorkbookWithSingleToken(workbook, fileId);

            // 3. Populate Data & Styles
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

            // Protect with password from properties
            sheet.protectSheet(sheetPassword);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private void signWorkbookWithSingleToken(XSSFWorkbook workbook, String fileId) {
        try {
            String signature = calculateHMAC(fileId, secretKey);
            String singleToken = fileId + "." + signature;

            POIXMLProperties props = workbook.getProperties();
            POIXMLProperties.CustomProperties customProps = props.getCustomProperties();
            if (customProps != null) {
                customProps.addProperty("Platform-Key", singleToken);
            }
            log.info("Generated Token: {}", singleToken);

        } catch (Exception e) {
            throw new RuntimeException("Could not sign file");
        }
    }

    public String verifyFileIntegrity(XSSFWorkbook workbook) {
        try {
            POIXMLProperties.CustomProperties props = workbook.getProperties().getCustomProperties();

            CTProperty tokenProp = props.getProperty("Platform-Key");
            if (tokenProp == null) throw new SecurityException("No File Key found.");

            String token = tokenProp.getLpwstr();
            String[] parts = token.split("\\.");
            if (parts.length != 2) throw new SecurityException("Invalid Key Format");

            String fileId = parts[0];
            String storedSig = parts[1];

            String calculatedSig = calculateHMAC(fileId, secretKey);

            if (!calculatedSig.equals(storedSig)) {
                throw new SecurityException("Signature Mismatch");
            }

            return fileId;

        } catch (SecurityException se) {
            throw se;
        } catch (Exception e) {
            throw new RuntimeException("Error verifying file");
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
        return style;
    }
}