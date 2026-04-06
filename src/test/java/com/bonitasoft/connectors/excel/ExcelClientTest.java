package com.bonitasoft.connectors.excel;

import static org.assertj.core.api.Assertions.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

class ExcelClientTest {

    @Test
    void shouldReadSheetData() throws Exception {
        byte[] content = createTestWorkbook("Sheet1",
                Arrays.asList("Name", "Age"),
                Arrays.asList("Alice", 30),
                Arrays.asList("Bob", 25));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .evaluateFormulas(true)
                .includeHeaders(true)
                .build();
        ExcelClient client = new ExcelClient(config);

        List<List<Object>> data = client.readSheet(content);

        assertThat(data).hasSize(3);
        assertThat(data.get(0)).containsExactly("Name", "Age");
        assertThat(data.get(1)).containsExactly("Alice", 30);
        assertThat(data.get(2)).containsExactly("Bob", 25);
    }

    @Test
    void shouldReadSheetWithoutHeaders() throws Exception {
        byte[] content = createTestWorkbook("Sheet1",
                Arrays.asList("Name", "Age"),
                Arrays.asList("Alice", 30));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .evaluateFormulas(true)
                .includeHeaders(false)
                .build();
        ExcelClient client = new ExcelClient(config);

        List<List<Object>> data = client.readSheet(content);

        assertThat(data).hasSize(1);
        assertThat(data.get(0)).containsExactly("Alice", 30);
    }

    @Test
    void shouldReadSheetByName() throws Exception {
        byte[] content = createTestWorkbook("MySheet",
                Arrays.asList("Col1"),
                Arrays.asList("Value1"));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetName("MySheet")
                .evaluateFormulas(true)
                .includeHeaders(true)
                .build();
        ExcelClient client = new ExcelClient(config);

        List<List<Object>> data = client.readSheet(content);

        assertThat(data).hasSize(2);
    }

    @Test
    void shouldThrowWhenSheetNotFound() throws Exception {
        byte[] content = createTestWorkbook("Sheet1", Arrays.asList("A"));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetName("NonExistent")
                .build();
        ExcelClient client = new ExcelClient(config);

        assertThatThrownBy(() -> client.readSheet(content))
                .isInstanceOf(ExcelException.class)
                .hasMessageContaining("Sheet not found");
    }

    @Test
    void shouldReadRange() throws Exception {
        byte[] content = createTestWorkbook("Sheet1",
                Arrays.asList("A", "B", "C"),
                Arrays.asList(1, 2, 3),
                Arrays.asList(4, 5, 6));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .range("B1:C2")
                .evaluateFormulas(true)
                .build();
        ExcelClient client = new ExcelClient(config);

        List<List<Object>> data = client.readRange(content);

        assertThat(data).hasSize(2);
        assertThat(data.get(0)).containsExactly("B", "C");
        assertThat(data.get(1)).containsExactly(2, 3);
    }

    @Test
    void shouldWriteSheet() throws Exception {
        byte[] content = createTestWorkbook("Sheet1", Arrays.asList("Old"));

        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("Name", "Score"));
        data.add(Arrays.asList("Alice", 95));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .data(data)
                .outputFileName("result.xlsx")
                .createSheetIfMissing(true)
                .streamingThreshold(10000)
                .build();
        ExcelClient client = new ExcelClient(config);

        byte[] result = client.writeSheet(content);
        assertThat(result).isNotEmpty();

        // Verify by reading back
        ExcelConfiguration readConfig = ExcelConfiguration.builder()
                .sheetIndex(0)
                .evaluateFormulas(true)
                .includeHeaders(true)
                .build();
        ExcelClient readClient = new ExcelClient(readConfig);
        List<List<Object>> readBack = readClient.readSheet(result);

        assertThat(readBack).hasSize(2);
        assertThat(readBack.get(0)).containsExactly("Name", "Score");
        assertThat(readBack.get(1)).containsExactly("Alice", 95);
    }

    @Test
    void shouldWriteRange() throws Exception {
        byte[] content = createTestWorkbook("Sheet1",
                Arrays.asList("A", "B", "C"),
                Arrays.asList(0, 0, 0),
                Arrays.asList(0, 0, 0));

        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList(10, 20));
        data.add(Arrays.asList(30, 40));

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .range("B2:C3")
                .data(data)
                .outputFileName("result.xlsx")
                .createSheetIfMissing(true)
                .build();
        ExcelClient client = new ExcelClient(config);

        byte[] result = client.writeRange(content);
        assertThat(result).isNotEmpty();

        // Verify range was written correctly
        ExcelConfiguration readConfig = ExcelConfiguration.builder()
                .sheetIndex(0)
                .range("B2:C3")
                .evaluateFormulas(true)
                .build();
        ExcelClient readClient = new ExcelClient(readConfig);
        List<List<Object>> readBack = readClient.readRange(result);

        assertThat(readBack).hasSize(2);
        assertThat(readBack.get(0)).containsExactly(10, 20);
        assertThat(readBack.get(1)).containsExactly(30, 40);
    }

    @Test
    void shouldThrowWhenWriteDataIsEmpty() {
        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .data(new ArrayList<>())
                .build();
        ExcelClient client = new ExcelClient(config);

        assertThatThrownBy(() -> client.writeSheet(new byte[0]))
                .isInstanceOf(ExcelException.class)
                .hasMessageContaining("No data provided");
    }

    @Test
    void shouldWriteSheetStreamingForLargeData() throws Exception {
        byte[] content = createTestWorkbook("Sheet1", Arrays.asList("Old"));

        List<List<Object>> data = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            data.add(Arrays.asList("Row" + i, i));
        }

        ExcelConfiguration config = ExcelConfiguration.builder()
                .sheetIndex(0)
                .data(data)
                .streamingThreshold(50) // Force streaming mode
                .build();
        ExcelClient client = new ExcelClient(config);

        byte[] result = client.writeSheet(content);
        assertThat(result).isNotEmpty();
    }

    // === Helper ===

    @SafeVarargs
    private byte[] createTestWorkbook(String sheetName, List<Object>... rows) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(sheetName);
            for (int i = 0; i < rows.length; i++) {
                var row = sheet.createRow(i);
                List<Object> rowData = rows[i];
                for (int j = 0; j < rowData.size(); j++) {
                    var cell = row.createCell(j);
                    CellValueConverter.writeCellValue(cell, rowData.get(j));
                }
            }
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        }
    }
}
