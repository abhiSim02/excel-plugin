package com.poc.excelplugin.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class RowData {
    /**
     * The row number in the original uploaded file (e.g., Row 5).
     * Useful for logging or telling the user exactly where the error is.
     */
    private int originalRowIndex;

    /**
     * Key: Column Index (0, 1, 2...)
     * Value: The actual cell value (String, Double, Boolean, LocalDateTime, etc.)
     */
    private Map<Integer, Object> cellValues;

    /**
     * If this is an Invalid Row (RED), this string contains the reason.
     * e.g., "[City: Invalid Value 'Mars']"
     * For Modified Rows (ORANGE), this is usually null.
     */
    private String errorMessage;
}