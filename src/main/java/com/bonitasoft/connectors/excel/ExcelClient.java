package com.bonitasoft.connectors.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel processing facade using Apache POI.
 * Handles reading and writing Excel files (.xlsx) with support for
 * formula evaluation and streaming writes for large datasets.
 */
@Slf4j
public class ExcelClient {

    private final ExcelConfiguration configuration;

    public ExcelClient(ExcelConfiguration configuration) {
        this.configuration = configuration;
        log.debug("ExcelClient initialized");
    }

    /**
     * Reads all data from a sheet.
     *
     * @param fileContent the Excel file bytes
     * @return list of rows, each row is a list of cell values
     */
    public List<List<Object>> readSheet(byte[] fileContent) throws ExcelException {
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(fileContent))) {
            Sheet sheet = resolveSheet(workbook);
            FormulaEvaluator evaluator = configuration.isEvaluateFormulas()
                    ? workbook.getCreationHelper().createFormulaEvaluator() : null;

            List<List<Object>> result = new ArrayList<>();
            int firstRow = configuration.isIncludeHeaders() ? sheet.getFirstRowNum() : sheet.getFirstRowNum() + 1;

            for (int rowIdx = firstRow; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) {
                    result.add(new ArrayList<>());
                    continue;
                }
                List<Object> rowData = new ArrayList<>();
                for (int colIdx = 0; colIdx < row.getLastCellNum(); colIdx++) {
                    Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    rowData.add(CellValueConverter.readCellValue(cell, evaluator));
                }
                result.add(rowData);
            }

            log.info("Read {} rows from sheet", result.size());
            return result;
        } catch (IOException e) {
            throw new ExcelException("Failed to read Excel file: " + e.getMessage(), e);
        }
    }

    /**
     * Reads data from a specific cell range (e.g., "A1:D10").
     *
     * @param fileContent the Excel file bytes
     * @return list of rows within the specified range
     */
    public List<List<Object>> readRange(byte[] fileContent) throws ExcelException {
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(fileContent))) {
            Sheet sheet = resolveSheet(workbook);
            FormulaEvaluator evaluator = configuration.isEvaluateFormulas()
                    ? workbook.getCreationHelper().createFormulaEvaluator() : null;

            CellRangeAddress rangeAddress = CellRangeAddress.valueOf(configuration.getRange());

            List<List<Object>> result = new ArrayList<>();
            for (int rowIdx = rangeAddress.getFirstRow(); rowIdx <= rangeAddress.getLastRow(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                List<Object> rowData = new ArrayList<>();
                for (int colIdx = rangeAddress.getFirstColumn(); colIdx <= rangeAddress.getLastColumn(); colIdx++) {
                    Cell cell = (row != null)
                            ? row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) : null;
                    rowData.add(CellValueConverter.readCellValue(cell, evaluator));
                }
                result.add(rowData);
            }

            log.info("Read {} rows from range {}", result.size(), configuration.getRange());
            return result;
        } catch (IOException e) {
            throw new ExcelException("Failed to read Excel range: " + e.getMessage(), e);
        }
    }

    /**
     * Writes data to a full sheet, replacing existing content.
     * Uses SXSSFWorkbook for streaming if data exceeds the streaming threshold.
     *
     * @param fileContent existing Excel file bytes (null to create new)
     * @return the resulting Excel file bytes
     */
    public byte[] writeSheet(byte[] fileContent) throws ExcelException {
        List<List<Object>> data = configuration.getData();
        if (data == null || data.isEmpty()) {
            throw new ExcelException("No data provided for write operation");
        }

        boolean useStreaming = data.size() > configuration.getStreamingThreshold();

        try {
            if (useStreaming) {
                return writeSheetStreaming(data);
            } else {
                return writeSheetStandard(fileContent, data);
            }
        } catch (IOException e) {
            throw new ExcelException("Failed to write Excel file: " + e.getMessage(), e);
        }
    }

    /**
     * Writes data to a specific cell range.
     *
     * @param fileContent existing Excel file bytes (null to create new)
     * @return the resulting Excel file bytes
     */
    public byte[] writeRange(byte[] fileContent) throws ExcelException {
        List<List<Object>> data = configuration.getData();
        if (data == null || data.isEmpty()) {
            throw new ExcelException("No data provided for write operation");
        }

        try (XSSFWorkbook workbook = fileContent != null
                ? new XSSFWorkbook(new ByteArrayInputStream(fileContent))
                : new XSSFWorkbook()) {

            Sheet sheet = resolveOrCreateSheet(workbook);
            CellRangeAddress rangeAddress = CellRangeAddress.valueOf(configuration.getRange());

            int startRow = rangeAddress.getFirstRow();
            int startCol = rangeAddress.getFirstColumn();

            for (int i = 0; i < data.size(); i++) {
                List<Object> rowData = data.get(i);
                Row row = sheet.getRow(startRow + i);
                if (row == null) {
                    row = sheet.createRow(startRow + i);
                }
                for (int j = 0; j < rowData.size(); j++) {
                    Cell cell = row.getCell(startCol + j);
                    if (cell == null) {
                        cell = row.createCell(startCol + j);
                    }
                    CellValueConverter.writeCellValue(cell, rowData.get(j));
                }
            }

            log.info("Wrote {} rows to range {}", data.size(), configuration.getRange());
            return toByteArray(workbook);
        } catch (IOException e) {
            throw new ExcelException("Failed to write Excel range: " + e.getMessage(), e);
        }
    }

    // === Private helpers ===

    private Sheet resolveSheet(Workbook workbook) throws ExcelException {
        Sheet sheet;
        if (configuration.getSheetName() != null && !configuration.getSheetName().isBlank()) {
            sheet = workbook.getSheet(configuration.getSheetName());
            if (sheet == null) {
                throw new ExcelException("Sheet not found: " + configuration.getSheetName());
            }
        } else {
            int index = configuration.getSheetIndex();
            if (index < 0 || index >= workbook.getNumberOfSheets()) {
                throw new ExcelException("Sheet index out of range: " + index);
            }
            sheet = workbook.getSheetAt(index);
        }
        return sheet;
    }

    private Sheet resolveOrCreateSheet(Workbook workbook) {
        if (configuration.getSheetName() != null && !configuration.getSheetName().isBlank()) {
            Sheet sheet = workbook.getSheet(configuration.getSheetName());
            if (sheet == null && configuration.isCreateSheetIfMissing()) {
                sheet = workbook.createSheet(configuration.getSheetName());
            }
            return sheet != null ? sheet : workbook.getSheetAt(0);
        }
        int index = configuration.getSheetIndex();
        if (index >= 0 && index < workbook.getNumberOfSheets()) {
            return workbook.getSheetAt(index);
        }
        if (workbook.getNumberOfSheets() == 0 && configuration.isCreateSheetIfMissing()) {
            return workbook.createSheet("Sheet1");
        }
        return workbook.getSheetAt(0);
    }

    private byte[] writeSheetStandard(byte[] fileContent, List<List<Object>> data) throws IOException {
        try (XSSFWorkbook workbook = fileContent != null
                ? new XSSFWorkbook(new ByteArrayInputStream(fileContent))
                : new XSSFWorkbook()) {

            Sheet sheet = resolveOrCreateSheet(workbook);
            // Clear existing content
            for (int i = sheet.getLastRowNum(); i >= sheet.getFirstRowNum(); i--) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }

            writeDataToSheet(sheet, data);
            log.info("Wrote {} rows to sheet (standard mode)", data.size());
            return toByteArray(workbook);
        }
    }

    private byte[] writeSheetStreaming(List<List<Object>> data) throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            Sheet sheet = workbook.createSheet(
                    configuration.getSheetName() != null && !configuration.getSheetName().isBlank()
                            ? configuration.getSheetName() : "Sheet1");

            writeDataToSheet(sheet, data);
            log.info("Wrote {} rows to sheet (streaming mode)", data.size());
            byte[] result = toByteArray(workbook);
            workbook.dispose();
            return result;
        }
    }

    private void writeDataToSheet(Sheet sheet, List<List<Object>> data) {
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            List<Object> rowData = data.get(i);
            for (int j = 0; j < rowData.size(); j++) {
                Cell cell = row.createCell(j);
                CellValueConverter.writeCellValue(cell, rowData.get(j));
            }
        }
    }

    private byte[] toByteArray(Workbook workbook) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        return baos.toByteArray();
    }
}
