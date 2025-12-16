package com.poc.excelplugin.service;

import com.poc.excelplugin.dto.AnalysisResult;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

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

    // Regex to find "_YYYYMMDDHHMMSS" at the end of a filename (before extension)
    private static final Pattern TIMESTAMP_PATTERN = Pattern.compile("(_\\d{14})$");

    public Map<String, Object> processUpload(MultipartFile file) throws Exception {

        // 1. Analyze the file
        AnalysisResult analysis = excelService.performDeepAnalysis(file);

        // 2. Generate new Timestamp
        String newTimestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));

        // --- CASE 1: RED (Validation Failed) ---
        if (!analysis.getInvalidRows().isEmpty()) {
            log.warn("Upload Rejected: Found {} invalid rows.", analysis.getInvalidRows().size());

            byte[] errorFile = excelService.generateSubsetExcel(analysis.getHeaders(), analysis.getInvalidRows(), true);

            // Format: ERROR_OriginalName_NewTimestamp.xlsx
            String errorFileName = generateCleanFileName(file.getOriginalFilename(), "ERROR_", newTimestamp);
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

            // Format: OriginalName_NewTimestamp.xlsx (No Prefix for Success)
            String deltaFileName = generateCleanFileName(file.getOriginalFilename(), "", newTimestamp);
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

    /**
     * Helper to clean filename, remove old timestamp, and add new one.
     * Input:  "MyFile_20251211104221.xlsx", prefix="ERROR_", newTime="20251216145859"
     * Output: "ERROR_MyFile_20251216145859.xlsx"
     */
    private String generateCleanFileName(String originalFilename, String prefix, String newTimestamp) {
        if (originalFilename == null) originalFilename = "Unknown_File.xlsx";

        // 1. Separate Name and Extension
        int dotIndex = originalFilename.lastIndexOf('.');
        String baseName = (dotIndex == -1) ? originalFilename : originalFilename.substring(0, dotIndex);
        String extension = (dotIndex == -1) ? ".xlsx" : originalFilename.substring(dotIndex);

        // 2. Strip OLD timestamp if it exists (looks for _14digits at end)
        Matcher matcher = TIMESTAMP_PATTERN.matcher(baseName);
        if (matcher.find()) {
            baseName = baseName.substring(0, matcher.start());
        }

        // 3. Build New Name
        return prefix + baseName + "_" + newTimestamp + extension;
    }
}