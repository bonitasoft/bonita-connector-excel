package com.bonitasoft.connectors.excel;

import lombok.Builder;
import lombok.Data;
import org.bonitasoft.engine.bpm.document.Document;

import java.util.List;

/**
 * Configuration for Excel connector operations.
 * Holds all parameters shared across read/write operations.
 */
@Data
@Builder
public class ExcelConfiguration {

    // === Document input (the Excel file from Bonita process) ===
    private Document document;

    // === Sheet identification ===
    private String sheetName;
    @Builder.Default
    private int sheetIndex = 0;

    // === Range parameters (for read-range / write-range) ===
    private String range;

    // === Read options ===
    @Builder.Default
    private boolean evaluateFormulas = true;
    @Builder.Default
    private boolean includeHeaders = true;

    // === Write options ===
    private List<List<Object>> data;
    @Builder.Default
    private String outputFileName = "output.xlsx";
    @Builder.Default
    private boolean createSheetIfMissing = true;
    @Builder.Default
    private int streamingThreshold = 10000;
}
