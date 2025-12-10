package com.poc.excelplugin.service;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@Service
public class LocalFileStorageService implements FileStorageService {

    @Value("${excel.storage.path:./output/}") // Default to ./output/ if not set
    private String storagePath;

    @Override
    public String saveFile(byte[] content, String fileName) throws IOException {
        Path path = Paths.get(storagePath + fileName);
        if (!Files.exists(path.getParent())) {
            Files.createDirectories(path.getParent());
        }
        Files.write(path, content);
        return path.toAbsolutePath().toString();
    }

    @Override
    public byte[] loadFile(String fileName) throws IOException {
        Path path = Paths.get(storagePath + fileName);
        return Files.readAllBytes(path);
    }
}