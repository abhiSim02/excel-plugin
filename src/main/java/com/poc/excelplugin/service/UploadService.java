package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.AnalysisResult;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
@Service
@RequiredArgsConstructor
public class UploadService {

    private final ExcelService excelService;
    private final FileStorageService fileStorageService;

    private static final Pattern TIMESTAMP_PATTERN = Pattern.compile("(_\\d{14})$");

    // UPDATED: Accepts InputStream and Filename String
    public Map<String, Object> processUpload(InputStream fileStream, String originalFilename) throws Exception {

        // 1. Analyze the file (Pass the stream)
        AnalysisResult analysis = excelService.performDeepAnalysis(fileStream);

        // 2. Generate new Timestamp
        String newTimestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));

        // --- CASE 1: RED (Validation Failed) ---
        if (!analysis.getInvalidRows().isEmpty()) {
            log.warn("Upload Rejected: Found {} invalid rows.", analysis.getInvalidRows().size());

            byte[] errorFile = excelService.generateSubsetExcel(analysis.getHeaders(), analysis.getInvalidRows(), true);

            String errorFileName = generateCleanFileName(originalFilename, "ERROR_", newTimestamp);
            String path = fileStorageService.saveFile(errorFile, "Validation_Failed/" + errorFileName);

            return Map.of(
                    "status", "FAILED",
                    "message", "Validation Errors Found.",
                    "errorCount", analysis.getInvalidRows().size(),
                    "reportPath", path
            );
        }

        // --- CASE 2: ORANGE (Modifications Found) ---
        if (!analysis.getModifiedRows().isEmpty()) {
            log.info("Upload Accepted: Found {} modified rows.", analysis.getModifiedRows().size());

            byte[] deltaFile = excelService.generateSubsetExcel(analysis.getHeaders(), analysis.getModifiedRows(), false);

            String deltaFileName = generateCleanFileName(originalFilename, "", newTimestamp);
            String path = fileStorageService.saveFile(deltaFile, "Inbound_Processing/" + deltaFileName);

            return Map.of(
                    "status", "SUCCESS",
                    "message", "Delta file generated.",
                    "modifiedCount", analysis.getModifiedRows().size(),
                    "processPath", path
            );
        }

        return Map.of("status", "IGNORED", "message", "No changes detected.");
    }

    private String generateCleanFileName(String originalFilename, String prefix, String newTimestamp) {
        if (originalFilename == null) originalFilename = "Unknown_File.xlsx";

        int dotIndex = originalFilename.lastIndexOf('.');
        String baseName = (dotIndex == -1) ? originalFilename : originalFilename.substring(0, dotIndex);
        String extension = (dotIndex == -1) ? ".xlsx" : originalFilename.substring(dotIndex);

        Matcher matcher = TIMESTAMP_PATTERN.matcher(baseName);
        if (matcher.find()) {
            baseName = baseName.substring(0, matcher.start());
        }

        return prefix + baseName + "_" + newTimestamp + extension;
    }
}