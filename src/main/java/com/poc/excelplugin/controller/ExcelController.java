package com.poc.excelplugin.controller;

import com.poc.excelplugin.dto.ApiResponse;
import com.poc.excelplugin.dto.ExcelRequest;
import com.poc.excelplugin.entity.UserFileHash;
import com.poc.excelplugin.repository.UserHashRepository;
import com.poc.excelplugin.service.ExcelService;
import com.poc.excelplugin.service.FileStorageService;
import com.poc.excelplugin.service.UploadService;
import jakarta.servlet.http.HttpServletRequest;
import lombok.RequiredArgsConstructor;
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

    private final ExcelService excelService;
    private final UserHashRepository userHashRepository;
    private final FileStorageService fileStorageService;
    private final UploadService uploadService;

    @PostMapping("/generate")
    public ResponseEntity<ApiResponse<String>> generateExcel(@RequestBody ExcelRequest request) {
        try {
            String timestamp = ZonedDateTime.now(ZoneId.of("UTC")).format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
            String fileName = request.getEntityName() + "_" + request.getUserId() + "_" + timestamp + ".xlsx";

            ExcelService.GenerationResult result = excelService.generateLargeExcel(request, timestamp);
            String storagePath = fileStorageService.saveFile(result.getFileContent(), fileName);

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

    /**
     * UPDATED: Accepts Raw Binary Stream
     * Usage in Postman:
     * 1. Method: POST
     * 2. Body -> Binary -> Select File
     * 3. Params -> Key: filename, Value: your_file.xlsx
     */
    @PostMapping(value = "/upload", consumes = "*/*")
    public ResponseEntity<ApiResponse<Map<String, Object>>> analyzeUpload(
            HttpServletRequest request,
            @RequestParam("filename") String filename
    ) {
        try {
            // Check content type to warn if user sends multipart by mistake
            String contentType = request.getContentType();
            if (contentType != null && contentType.toLowerCase().contains("multipart/form-data")) {
                return ResponseEntity.badRequest().body(ApiResponse.error("Incorrect Upload Method. Please select 'Binary' body in Postman, not 'form-data'.", "INVALID_REQ"));
            }

            // Pass the raw stream and the filename from the query param
            Map<String, Object> result = uploadService.processUpload(request.getInputStream(), filename);

            if ("FAILED".equals(result.get("status"))) {
                return ResponseEntity.badRequest().body(ApiResponse.error("Validation Failed. Error report generated.", result.get("reportPath").toString()));
            }

            return ResponseEntity.ok(ApiResponse.success("File Processed Successfully", result));
        } catch (Exception e) {
            log.error("Upload failed", e);
            return ResponseEntity.internalServerError().body(ApiResponse.error(e.getMessage(), "UPLOAD_FAIL"));
        }
    }
}