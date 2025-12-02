package com.poc.excelplugin.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import java.util.List;
import java.util.Map;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExcelRequest {
    private String entityName;
    private String userId;

    private List<ColumnConfig> columns;
    private List<Map<String, Object>> data;

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class ColumnConfig {
        private String header;
        private String key;
        private boolean editable;
        private List<String> dropdown;
        private Integer width;
        private String dataFormat;
    }
}