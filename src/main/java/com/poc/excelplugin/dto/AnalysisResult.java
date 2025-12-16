package com.poc.excelplugin.dto;

import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class AnalysisResult {
    /**
     * The header names extracted from the uploaded file.
     * Needed so we can recreate the Excel file structure when dumping errors/deltas.
     */
    private List<String> headers;

    /**
     * List of rows that failed validation (RED rows).
     * If this list is not empty, the whole upload should usually be rejected.
     */
    private List<RowData> invalidRows;

    /**
     * List of rows that are valid but have changed (ORANGE rows).
     * These are the candidate rows for database updates.
     */
    private List<RowData> modifiedRows;
}