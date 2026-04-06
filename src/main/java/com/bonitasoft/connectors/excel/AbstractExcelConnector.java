package com.bonitasoft.connectors.excel;

import lombok.extern.slf4j.Slf4j;
import org.bonitasoft.engine.bpm.document.Document;
import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;

import java.util.List;

/**
 * Abstract base connector for Excel operations.
 * Reads a Bonita Document containing an Excel file, processes it via Apache POI,
 * and optionally produces a new Document as output.
 */
@Slf4j
public abstract class AbstractExcelConnector extends AbstractConnector {

    // Common output parameter constants
    protected static final String OUTPUT_SUCCESS = "success";
    protected static final String OUTPUT_ERROR_MESSAGE = "errorMessage";

    protected ExcelConfiguration configuration;
    protected ExcelClient client;

    @Override
    public void validateInputParameters() throws ConnectorValidationException {
        try {
            this.configuration = buildConfiguration();
            validateConfiguration(this.configuration);
        } catch (IllegalArgumentException e) {
            throw new ConnectorValidationException(this, e.getMessage());
        }
    }

    @Override
    public void connect() throws ConnectorException {
        this.client = new ExcelClient(this.configuration);
        log.info("Excel connector initialized");
    }

    @Override
    public void disconnect() throws ConnectorException {
        this.client = null;
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            doExecute();
            setOutputParameter(OUTPUT_SUCCESS, true);
        } catch (ExcelException e) {
            log.error("Excel connector execution failed: {}", e.getMessage(), e);
            setOutputParameter(OUTPUT_SUCCESS, false);
            setOutputParameter(OUTPUT_ERROR_MESSAGE, e.getMessage());
        } catch (Exception e) {
            log.error("Unexpected error in Excel connector: {}", e.getMessage(), e);
            setOutputParameter(OUTPUT_SUCCESS, false);
            setOutputParameter(OUTPUT_ERROR_MESSAGE, "Unexpected error: " + e.getMessage());
        }
    }

    protected abstract void doExecute() throws ExcelException, ConnectorException;

    protected abstract ExcelConfiguration buildConfiguration();

    protected void validateConfiguration(ExcelConfiguration config) {
        if (config.getDocument() == null) {
            throw new IllegalArgumentException("document is mandatory");
        }
    }

    /**
     * Retrieves the file content bytes from a Bonita Document using the API accessor.
     */
    protected byte[] getDocumentContent(Document document) throws ConnectorException {
        try {
            return getAPIAccessor().getProcessAPI()
                    .getDocumentContent(document.getContentStorageId());
        } catch (Exception e) {
            throw new ConnectorException("Failed to retrieve document content: " + e.getMessage(), e);
        }
    }

    // === Input reading helpers ===

    protected String readStringInput(String name) {
        Object value = getInputParameter(name);
        return value != null ? value.toString() : null;
    }

    protected String readStringInput(String name, String defaultValue) {
        String value = readStringInput(name);
        return (value != null && !value.isBlank()) ? value : defaultValue;
    }

    protected Boolean readBooleanInput(String name, boolean defaultValue) {
        Object value = getInputParameter(name);
        return value != null ? (Boolean) value : defaultValue;
    }

    protected Integer readIntegerInput(String name, int defaultValue) {
        Object value = getInputParameter(name);
        return value != null ? ((Number) value).intValue() : defaultValue;
    }

    @SuppressWarnings("unchecked")
    protected List<List<Object>> readListOfListInput(String name) {
        Object value = getInputParameter(name);
        return value != null ? (List<List<Object>>) value : null;
    }
}
