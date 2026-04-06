package com.bonitasoft.connectors.excel;

import static org.assertj.core.api.Assertions.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class CellValueConverterTest {

    private Workbook workbook;
    private Sheet sheet;
    private Row row;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
        row = sheet.createRow(0);
    }

    @AfterEach
    void tearDown() throws Exception {
        workbook.close();
    }

    @Test
    void shouldReadStringCell() {
        Cell cell = row.createCell(0);
        cell.setCellValue("Hello");
        assertThat(CellValueConverter.readCellValue(cell, null)).isEqualTo("Hello");
    }

    @Test
    void shouldReadIntegerNumericCell() {
        Cell cell = row.createCell(0);
        cell.setCellValue(42.0);
        assertThat(CellValueConverter.readCellValue(cell, null)).isEqualTo(42);
    }

    @Test
    void shouldReadDoubleNumericCell() {
        Cell cell = row.createCell(0);
        cell.setCellValue(3.14);
        assertThat(CellValueConverter.readCellValue(cell, null)).isEqualTo(3.14);
    }

    @Test
    void shouldReadBooleanCell() {
        Cell cell = row.createCell(0);
        cell.setCellValue(true);
        assertThat(CellValueConverter.readCellValue(cell, null)).isEqualTo(true);
    }

    @Test
    void shouldReadBlankCellAsNull() {
        Cell cell = row.createCell(0);
        cell.setBlank();
        assertThat(CellValueConverter.readCellValue(cell, null)).isNull();
    }

    @Test
    void shouldReadNullCellAsNull() {
        assertThat(CellValueConverter.readCellValue(null, null)).isNull();
    }

    @Test
    void shouldReadFormulaWithEvaluator() {
        Cell cellA = row.createCell(0);
        cellA.setCellValue(10);
        Cell cellB = row.createCell(1);
        cellB.setCellValue(20);
        Cell cellC = row.createCell(2);
        cellC.setCellFormula("A1+B1");

        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Object result = CellValueConverter.readCellValue(cellC, evaluator);
        assertThat(result).isEqualTo(30);
    }

    @Test
    void shouldReadFormulaWithoutEvaluatorAsString() {
        Cell cell = row.createCell(0);
        cell.setCellFormula("A2+B2");
        Object result = CellValueConverter.readCellValue(cell, null);
        assertThat(result).isEqualTo("A2+B2");
    }

    @Test
    void shouldWriteStringValue() {
        Cell cell = row.createCell(0);
        CellValueConverter.writeCellValue(cell, "test");
        assertThat(cell.getStringCellValue()).isEqualTo("test");
    }

    @Test
    void shouldWriteNumericValue() {
        Cell cell = row.createCell(0);
        CellValueConverter.writeCellValue(cell, 42);
        assertThat(cell.getNumericCellValue()).isEqualTo(42.0);
    }

    @Test
    void shouldWriteBooleanValue() {
        Cell cell = row.createCell(0);
        CellValueConverter.writeCellValue(cell, true);
        assertThat(cell.getBooleanCellValue()).isTrue();
    }

    @Test
    void shouldWriteNullAsBlank() {
        Cell cell = row.createCell(0);
        CellValueConverter.writeCellValue(cell, null);
        assertThat(cell.getCellType()).isEqualTo(CellType.BLANK);
    }

    @Test
    void shouldWriteDoubleValue() {
        Cell cell = row.createCell(0);
        CellValueConverter.writeCellValue(cell, 3.14);
        assertThat(cell.getNumericCellValue()).isEqualTo(3.14);
    }

    @Test
    void shouldWriteObjectAsString() {
        Cell cell = row.createCell(0);
        CellValueConverter.writeCellValue(cell, new Object() {
            @Override
            public String toString() {
                return "custom";
            }
        });
        assertThat(cell.getStringCellValue()).isEqualTo("custom");
    }
}
