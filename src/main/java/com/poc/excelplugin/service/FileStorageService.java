package com.poc.excelplugin.service;

import java.io.IOException;

public interface FileStorageService {
    /**
     * Saves byte content to the storage system.
     * @param content The file data.
     * @param fileName The desired filename.
     * @return The path or URL where the file is stored.
     */
    String saveFile(byte[] content, String fileName) throws IOException;

    /**
     * Loads a file from storage.
     * @param fileName The name of the file to load.
     * @return The file data as a byte array.
     */
    byte[] loadFile(String fileName) throws IOException;
}