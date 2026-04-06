package com.bonitasoft.connectors.excel;

import lombok.extern.slf4j.Slf4j;
import org.bonitasoft.engine.bpm.document.Document;
import org.bonitasoft.engine.bpm.document.DocumentValue;
import org.bonitasoft.engine.connector.ConnectorException;

import java.util.List;

/**
 * Writes data to a full sheet in an Excel document.
 * Uses SXSSFWorkbook for streaming writes when data exceeds 10,000 rows.
 */
@Slf4j
public class WriteSheetConnector extends AbstractExcelConnector {

    // Input parameter names
    static final String INPUT_DOCUMENT = "document";
    static final String INPUT_SHEET_NAME = "sheetName";
    static final String INPUT_SHEET_INDEX = "sheetIndex";
    static final String INPUT_DATA = "data";
    static final String INPUT_OUTPUT_FILE_NAME = "outputFileName";
    static final String INPUT_CREATE_SHEET_IF_MISSING = "createSheetIfMissing";
    static final String INPUT_STREAMING_THRESHOLD = "streamingThreshold";

    // Output parameter names
    static final String OUTPUT_DOCUMENT_VALUE = "documentValue";
    static final String OUTPUT_ROW_COUNT = "rowCount";

    @Override
    protected ExcelConfiguration buildConfiguration() {
        return ExcelConfiguration.builder()
                .document((Document) getInputParameter(INPUT_DOCUMENT))
                .sheetName(readStringInput(INPUT_SHEET_NAME))
                .sheetIndex(readIntegerInput(INPUT_SHEET_INDEX, 0))
                .data(readListOfListInput(INPUT_DATA))
                .outputFileName(readStringInput(INPUT_OUTPUT_FILE_NAME, "output.xlsx"))
                .createSheetIfMissing(readBooleanInput(INPUT_CREATE_SHEET_IF_MISSING, true))
                .streamingThreshold(readIntegerInput(INPUT_STREAMING_THRESHOLD, 10000))
                .build();
    }

    @Override
    protected void validateConfiguration(ExcelConfiguration config) {
        super.validateConfiguration(config);
        if (config.getData() == null || config.getData().isEmpty()) {
            throw new IllegalArgumentException("data is mandatory and must not be empty");
        }
    }

    @Override
    protected void doExecute() throws ExcelException, ConnectorException {
        log.info("Executing WriteSheet connector");

        byte[] fileContent = getDocumentContent(configuration.getDocument());
        byte[] result = client.writeSheet(fileContent);

        DocumentValue documentValue = new DocumentValue(result,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                configuration.getOutputFileName());

        setOutputParameter(OUTPUT_DOCUMENT_VALUE, documentValue);
        setOutputParameter(OUTPUT_ROW_COUNT, configuration.getData().size());

        log.info("WriteSheet connector executed successfully: {} rows written", configuration.getData().size());
    }
}
