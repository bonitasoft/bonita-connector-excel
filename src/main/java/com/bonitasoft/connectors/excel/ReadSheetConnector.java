package com.bonitasoft.connectors.excel;

import lombok.extern.slf4j.Slf4j;
import org.bonitasoft.engine.bpm.document.Document;
import org.bonitasoft.engine.connector.ConnectorException;

import java.util.List;

/**
 * Reads all data from a sheet in an Excel document.
 */
@Slf4j
public class ReadSheetConnector extends AbstractExcelConnector {

    // Input parameter names
    static final String INPUT_DOCUMENT = "document";
    static final String INPUT_SHEET_NAME = "sheetName";
    static final String INPUT_SHEET_INDEX = "sheetIndex";
    static final String INPUT_EVALUATE_FORMULAS = "evaluateFormulas";
    static final String INPUT_INCLUDE_HEADERS = "includeHeaders";

    // Output parameter names
    static final String OUTPUT_DATA = "data";
    static final String OUTPUT_ROW_COUNT = "rowCount";

    @Override
    protected ExcelConfiguration buildConfiguration() {
        return ExcelConfiguration.builder()
                .document((Document) getInputParameter(INPUT_DOCUMENT))
                .sheetName(readStringInput(INPUT_SHEET_NAME))
                .sheetIndex(readIntegerInput(INPUT_SHEET_INDEX, 0))
                .evaluateFormulas(readBooleanInput(INPUT_EVALUATE_FORMULAS, true))
                .includeHeaders(readBooleanInput(INPUT_INCLUDE_HEADERS, true))
                .build();
    }

    @Override
    protected void doExecute() throws ExcelException, ConnectorException {
        log.info("Executing ReadSheet connector");

        byte[] fileContent = getDocumentContent(configuration.getDocument());
        List<List<Object>> data = client.readSheet(fileContent);

        setOutputParameter(OUTPUT_DATA, data);
        setOutputParameter(OUTPUT_ROW_COUNT, data.size());

        log.info("ReadSheet connector executed successfully: {} rows read", data.size());
    }
}
