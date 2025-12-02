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

            // 2. SAVE TO DB: Store [User ID] + [Hash Key]
            UserFileHash dbEntry = new UserFileHash();
            dbEntry.setUserId(user);
            dbEntry.setHashKey(signature);
            userHashRepository.save(dbEntry);

            log.info("DB SAVED: User [{}] -> Hash [{}]", user, signature);

            // 3. Save File to Disk
            String filename = entity + ".xlsx";
            Path path = Paths.get(outputFolder + filename);
            if (!Files.exists(path.getParent())) {
                Files.createDirectories(path.getParent());
            }
            Files.write(path, result.getFileContent());

            return ResponseEntity.ok("SUCCESS: File created & Hash stored in DB.");

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

            // 1. Extract the full metadata string from the file
            // Format: Entity|User|Time|Signature
            String fullToken = excelService.extractPlatformKey(workbook);
            String[] parts = fullToken.split("\\|");

            if (parts.length < 4) {
                return ResponseEntity.status(400).body("❌ INVALID FORMAT: Token missing components.");
            }

            String fileUser = parts[1];      // User ID from file metadata
            String fileSignature = parts[3]; // The 24-char Hash

            // 2. DB VERIFICATION: Check if this Hash exists for this User in Postgres
            boolean isValid = userHashRepository.existsByUserIdAndHashKey(fileUser, fileSignature);

            if (isValid) {
                // Optional: Verify Math (to ensure metadata wasn't tampered to match DB hash)
                excelService.verifyMathIntegrity(parts[0], parts[1], parts[2], parts[3]);

                return ResponseEntity.ok("✅ VALIDATED VIA DB.\nDatabase confirms this signature belongs to user: " + fileUser);
            } else {
                log.warn("DB Lookup Failed. User: {}, Signature: {}", fileUser, fileSignature);
                return ResponseEntity.status(401).body("❌ UNAUTHORIZED: This file's signature was not found in the Database.");
            }

        } catch (SecurityException e) {
            return ResponseEntity.status(401).body("❌ TAMPERED: " + e.getMessage());
        } catch (IOException e) {
            return ResponseEntity.internalServerError().body("Error reading file: " + e.getMessage());
        }
    }
}