package com.poc.excelplugin.controller;

import com.poc.excelplugin.dto.ExcelRequest;
import com.poc.excelplugin.entity.UserFileHash;
import com.poc.excelplugin.repository.UserHashRepository;
import com.poc.excelplugin.service.ExcelService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Optional;

@Slf4j
@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @Autowired
    private UserHashRepository userHashRepository;

    @Value("${excel.storage.path}")
    private String outputFolder;

    @PostMapping("/generate")
    public ResponseEntity<String> generateAndSaveExcel(@RequestBody ExcelRequest request) {
        String entity = (request.getEntityName() != null) ? request.getEntityName() : "output";
        String user = (request.getUserId() != null) ? request.getUserId() : "unknown_user";

        log.info(">>> API HIT: Generate request for Entity: '{}', User: '{}'", entity, user);

        try {
            // 1. Generate File AND get the 24-char Signature
            ExcelService.GenerationResult result = excelService.generateGenericExcel(request);
            String signature = result.getSignature();

            // 2. DB LOGIC: Check if record exists for this User + Entity
            Optional<UserFileHash> existingRecord = userHashRepository.findByUserIdAndEntityName(user, entity);

            UserFileHash dbEntry;
            if (existingRecord.isPresent()) {
                // UPDATE Existing
                dbEntry = existingRecord.get();
                dbEntry.setHashKey(signature); // Overwrite old hash
                log.info("DB UPDATE: Found existing record for User [{}] Entity [{}]. Updating Hash.", user, entity);
            } else {
                // CREATE New
                dbEntry = new UserFileHash();
                dbEntry.setUserId(user);
                dbEntry.setEntityName(entity);
                dbEntry.setHashKey(signature);
                log.info("DB INSERT: Creating new record for User [{}] Entity [{}].", user, entity);
            }

            // Save (Timestamps handled automatically by @PrePersist/@PreUpdate)
            userHashRepository.save(dbEntry);

            // 3. Save File to Disk
            String filename = entity + ".xlsx";
            Path path = Paths.get(outputFolder + filename);
            if (!Files.exists(path.getParent())) {
                Files.createDirectories(path.getParent());
            }
            Files.write(path, result.getFileContent());

            return ResponseEntity.ok("SUCCESS: File created. Database updated for Entity: " + entity);

        } catch (IOException e) {
            log.error(">>> ERROR: IO Exception.", e);
            return ResponseEntity.internalServerError().body("ERROR: " + e.getMessage());
        } catch (Exception e) {
            log.error(">>> ERROR: Unexpected error.", e);
            return ResponseEntity.internalServerError().body("Unexpected Error: " + e.getMessage());
        }
    }

    @PostMapping("/verify")
    public ResponseEntity<String> verifyExcelFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().body("Please upload a file.");
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream())) {

            String fullToken = excelService.extractPlatformKey(workbook);
            String[] parts = fullToken.split("\\|");

            if (parts.length < 4) {
                return ResponseEntity.status(400).body("❌ INVALID FORMAT: Token missing components.");
            }

            String fileEntity = parts[0];
            String fileUser = parts[1];
            String fileSignature = parts[3];

            // DB VERIFICATION
            // We verify if this specific Hash exists.
            // Note: If the user generated a NEW file for this entity, the old hash is gone from DB,
            // so the old file will correctly fail verification.
            boolean isValid = userHashRepository.existsByUserIdAndHashKey(fileUser, fileSignature);

            if (isValid) {
                return ResponseEntity.ok("✅ VALIDATED VIA DB.\nDatabase confirms this is the latest valid file for User: " + fileUser);
            } else {
                return ResponseEntity.status(401).body("❌ UNAUTHORIZED: Hash not found. (User may have generated a newer version of this file).");
            }

        } catch (SecurityException e) {
            return ResponseEntity.status(401).body("❌ TAMPERED: " + e.getMessage());
        } catch (IOException e) {
            return ResponseEntity.internalServerError().body("Error reading file: " + e.getMessage());
        }
    }
}