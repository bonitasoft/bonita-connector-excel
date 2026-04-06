package com.bonitasoft.connectors.excel;

import lombok.extern.slf4j.Slf4j;
import org.bonitasoft.engine.bpm.document.Document;
import org.bonitasoft.engine.connector.ConnectorException;

import java.util.List;

/**
 * Reads data from a specific cell range in an Excel document.
 */
@Slf4j
public class ReadRangeConnector extends AbstractExcelConnector {

    // Input parameter names
    static final String INPUT_DOCUMENT = "document";
    static final String INPUT_SHEET_NAME = "sheetName";
    static final String INPUT_SHEET_INDEX = "sheetIndex";
    static final String INPUT_RANGE = "range";
    static final String INPUT_EVALUATE_FORMULAS = "evaluateFormulas";

    // Output parameter names
    static final String OUTPUT_DATA = "data";
    static final String OUTPUT_ROW_COUNT = "rowCount";

    @Override
    protected ExcelConfiguration buildConfiguration() {
        return ExcelConfiguration.builder()
                .document((Document) getInputParameter(INPUT_DOCUMENT))
                .sheetName(readStringInput(INPUT_SHEET_NAME))
                .sheetIndex(readIntegerInput(INPUT_SHEET_INDEX, 0))
                .range(readStringInput(INPUT_RANGE))
                .evaluateFormulas(readBooleanInput(INPUT_EVALUATE_FORMULAS, true))
                .build();
    }

    @Override
    protected void validateConfiguration(ExcelConfiguration config) {
        super.validateConfiguration(config);
        if (config.getRange() == null || config.getRange().isBlank()) {
            throw new IllegalArgumentException("range is mandatory (e.g., A1:D10)");
        }
    }

    @Override
    protected void doExecute() throws ExcelException, ConnectorException {
        log.info("Executing ReadRange connector for range: {}", configuration.getRange());

        byte[] fileContent = getDocumentContent(configuration.getDocument());
        List<List<Object>> data = client.readRange(fileContent);

        setOutputParameter(OUTPUT_DATA, data);
        setOutputParameter(OUTPUT_ROW_COUNT, data.size());

        log.info("ReadRange connector executed successfully: {} rows read", data.size());
    }
}
