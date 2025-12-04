//package com.sim.spriced.platform.excel_plugin.service;
//
//import com.sim.spriced.platform.data_entity_orchestrator.data_persistance.service.GenericQueryService;
//import com.sim.spriced.platform.data_entity_orchestrator.entity_handler.Attribute.BaseAttribute;
//import com.sim.spriced.platform.data_entity_orchestrator.entity_handler.Entities.BaseEntity;
//import com.sim.spriced.platform.data_entity_orchestrator.entity_handler.EntityManager;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddressList;
//import org.apache.poi.xssf.usermodel.*;
//import org.jooq.Record;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.stereotype.Service;
//
//import java.io.ByteArrayOutputStream;
//import java.util.*;
//import java.util.stream.Collectors;
//import java.util.stream.Stream;
//
//@Service
//public class ExcelGenerationService {
//
//    private static final Logger logger = LoggerFactory.getLogger(ExcelGenerationService.class);
//
//    @Autowired
//    private EntityManager entityManager;
//
//    public byte[] generateSecureExcel(String entityName, String userId) {
//        logger.info("Generating Excel (POC Match) for Entity: {}", entityName);
//
//        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
//
//            // 1. Fetch Schema
//            BaseEntity entityDef = entityManager.get(entityName);
//            if (entityDef == null) throw new RuntimeException("Entity not found: " + entityName);
//
//            // Sort Attributes (Required for alignment)
//            List<BaseAttribute> orderedAttributes = new ArrayList<>(entityDef.getAttribute().values());
//            orderedAttributes.sort(Comparator.comparing(BaseAttribute::getName));
//
//            // 2. Fetch Data
//            String sql = "select * from " + entityName;
//            List<Map<String, Object>> dataRows = new ArrayList<>();
//            try (Stream<Record> recordStream = GenericQueryService.executeQueryForStream(sql, entityManager.getDsl())) {
//                if (recordStream != null) {
//                    dataRows = recordStream.map(Record::intoMap).collect(Collectors.toList());
//                }
//            } catch (Exception e) {
//                logger.warn("Data fetch failed: {}", e.getMessage());
//            }
//
//            // 3. Create Sheet
//            XSSFSheet sheet = workbook.createSheet(entityName);
//            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
//
//            // 4. Create Styles
//            CellStyle headerStyle = workbook.createCellStyle();
//            Font font = workbook.createFont();
//            font.setBold(true);
//            headerStyle.setFont(font);
//            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//            // 5. Fill Headers
//            Row header = sheet.createRow(0);
//            int col = 0;
//            for (BaseAttribute attr : orderedAttributes) {
//                Cell cell = header.createCell(col++);
//                cell.setCellValue(attr.getName());
//                cell.setCellStyle(headerStyle);
//                sheet.setColumnWidth(col - 1, 6000);
//            }
//
//            // 6. Fill Data
//            int rowIdx = 1;
//            for (Map<String, Object> rowMap : dataRows) {
//                Row row = sheet.createRow(rowIdx++);
//                col = 0;
//                for (BaseAttribute attr : orderedAttributes) {
//                    Cell cell = row.createCell(col++);
//                    Object val = rowMap.get(attr.getName());
//                    if (val != null) {
//                        if (val instanceof Number) cell.setCellValue(((Number) val).doubleValue());
//                        else cell.setCellValue(val.toString());
//                    }
//                }
//            }
//
//            // 7. ADD DROPDOWNS (Exactly like POC)
//            int colIndex = 0;
//            for (BaseAttribute attr : orderedAttributes) {
//                List<String> validValues = attr.getValidValues();
//
//                if (isValidDropdownList(validValues)) {
//                    // Clean List
//                    String[] options = validValues.stream()
//                            .filter(s -> s != null && !s.trim().isEmpty())
//                            .toArray(String[]::new);
//
//                    if (options.length > 0) {
//                        logger.info("Adding Dropdown for '{}': {}", attr.getName(), Arrays.toString(options));
//
//                        // Create Explicit List Constraint
//                        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
//
//                        // Apply to Row 1 (Index 1) to Row 5000
//                        CellRangeAddressList addressList = new CellRangeAddressList(1, 5000, colIndex, colIndex);
//
//                        DataValidation validation = validationHelper.createValidation(constraint, addressList);
//
//                        // --- MATCHING POC EXACTLY ---
//                        validation.setShowErrorBox(true);
//                        // REMOVED: setSuppressDropDownArrow(false) -> Rely on default
//                        // REMOVED: setEmptyCellAllowed(true) -> Rely on default
//
//                        sheet.addValidationData(validation);
//                    }
//                }
//                colIndex++;
//            }
//
//            // 8. Write File
//            ByteArrayOutputStream out = new ByteArrayOutputStream();
//            workbook.write(out);
//            return out.toByteArray();
//
//        } catch (Exception e) {
//            throw new RuntimeException("Excel generation failed: " + e.getMessage());
//        }
//    }
//
//    private boolean isValidDropdownList(List<String> list) {
//        if (list == null || list.isEmpty()) return false;
//        return list.stream().anyMatch(s -> s != null && !s.trim().isEmpty());
//    }
//}