package com.poc.excelplugin.controller;

import com.poc.excelplugin.dto.ApiResponse;
import com.poc.excelplugin.dto.ExcelRequest;
import com.poc.excelplugin.entity.UserFileHash;
import com.poc.excelplugin.repository.UserHashRepository;
import com.poc.excelplugin.service.ExcelService;
import com.poc.excelplugin.service.FileStorageService;
import com.poc.excelplugin.service.UploadService;
import lombok.RequiredArgsConstructor; // <--- IMP: Import this
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Map;
import java.util.Optional;

@Slf4j
@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelController {

    // Make all services FINAL so Lombok injects them automatically
    private final ExcelService excelService;
    private final UserHashRepository userHashRepository;
    private final FileStorageService fileStorageService;
    private final UploadService uploadService;

    @PostMapping("/generate")
    public ResponseEntity<ApiResponse<String>> generateExcel(@RequestBody ExcelRequest request) {
        try {
            // 1. UTC Timestamp
            String timestamp = ZonedDateTime.now(ZoneId.of("UTC")).format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
            String fileName = request.getEntityName() + "_" + request.getUserId() + "_" + timestamp + ".xlsx";

            // 2. Generate (SXSSF Streaming)
            ExcelService.GenerationResult result = excelService.generateLargeExcel(request, timestamp);

            // 3. Save to Storage
            String storagePath = fileStorageService.saveFile(result.getFileContent(), fileName);

            // 4. Update DB Hash (Overwrites previous active file for this user/entity)
            Optional<UserFileHash> existing = userHashRepository.findByUserIdAndEntityName(request.getUserId(), request.getEntityName());
            UserFileHash hashEntry = existing.orElse(new UserFileHash());
            hashEntry.setUserId(request.getUserId());
            hashEntry.setEntityName(request.getEntityName());
            hashEntry.setHashKey(result.getSignature());
            userHashRepository.save(hashEntry);

            return ResponseEntity.ok(ApiResponse.success("File Generated Successfully", storagePath));

        } catch (Exception e) {
            log.error("Generation error", e);
            return ResponseEntity.internalServerError().body(ApiResponse.error(e.getMessage(), "GEN_FAIL"));
        }
    }

    @PostMapping("/verify")
    public ResponseEntity<ApiResponse<String>> verifyExcel(@RequestParam("file") MultipartFile file) {
        return ResponseEntity.ok(ApiResponse.success("Verification logic placeholder", "Use SAX Reader for 1M rows"));
    }

    @PostMapping("/upload")
    public ResponseEntity<ApiResponse<Map<String, Object>>> analyzeUpload(@RequestParam("file") MultipartFile file) {
        try {
            // The Service decides if it's a "Success" (Delta generated) or "Failure" (Error dump generated)
            Map<String, Object> result = uploadService.processUpload(file);

            if ("FAILED".equals(result.get("status"))) {
                // Return 400 Bad Request if validation failed, but include the path to the Error Report
                // This allows the frontend to show "Validation Failed" and provide a "Download Error Report" button
                return ResponseEntity.badRequest().body(ApiResponse.error("Validation Failed. Error report generated.", result.get("reportPath").toString()));
            }

            return ResponseEntity.ok(ApiResponse.success("File Processed Successfully", result));
        } catch (Exception e) {
            log.error("Upload failed", e);
            return ResponseEntity.internalServerError().body(ApiResponse.error(e.getMessage(), "UPLOAD_FAIL"));
        }
    }
}