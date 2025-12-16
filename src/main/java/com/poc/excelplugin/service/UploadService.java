package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.AnalysisResult;
import com.poc.excelplugin.dto.RowData;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class UploadService {

    private final ExcelService excelService;
    private final FileStorageService fileStorageService;

    public Map<String, Object> processUpload(MultipartFile file) throws Exception {

        // 1. Analyze the file (Get all Red/Orange/Clean rows)
        AnalysisResult analysis = excelService.performDeepAnalysis(file);

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        // --- CASE 1: RED (Validation Failed) ---
        if (!analysis.getInvalidRows().isEmpty()) {
            log.warn("Upload Rejected: Found {} invalid rows.", analysis.getInvalidRows().size());

            // Logic: Dump ONLY Red rows to "Validation_Failed" folder
            byte[] errorFile = excelService.generateSubsetExcel(analysis.getHeaders(), analysis.getInvalidRows(), true);

            String errorFileName = "ERROR_" + timestamp + "_" + file.getOriginalFilename();
            String path = fileStorageService.saveFile(errorFile, "Validation_Failed/" + errorFileName);

            return Map.of(
                    "status", "FAILED",
                    "message", "Validation Errors Found. Processing Aborted.",
                    "errorCount", analysis.getInvalidRows().size(),
                    "reportPath", path
            );
        }

        // --- CASE 2: ORANGE (Modifications Found) ---
        if (!analysis.getModifiedRows().isEmpty()) {
            log.info("Upload Accepted: Found {} modified rows.", analysis.getModifiedRows().size());

            // Logic: Dump ONLY Orange rows to "Inbound" folder (for downstream processing)
            byte[] deltaFile = excelService.generateSubsetExcel(analysis.getHeaders(), analysis.getModifiedRows(), false);

            String deltaFileName = "DELTA_" + timestamp + "_" + file.getOriginalFilename();
            String path = fileStorageService.saveFile(deltaFile, "Inbound_Processing/" + deltaFileName);

            return Map.of(
                    "status", "SUCCESS",
                    "message", "Delta file generated for processing.",
                    "modifiedCount", analysis.getModifiedRows().size(),
                    "processPath", path
            );
        }

        // --- CASE 3: NO CHANGES ---
        return Map.of("status", "IGNORED", "message", "No changes detected.");
    }
}