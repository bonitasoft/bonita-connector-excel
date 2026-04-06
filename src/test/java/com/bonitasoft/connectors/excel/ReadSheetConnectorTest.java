package com.bonitasoft.connectors.excel;

import static org.assertj.core.api.Assertions.*;
import static org.mockito.Mockito.*;

import org.bonitasoft.engine.api.APIAccessor;
import org.bonitasoft.engine.api.ProcessAPI;
import org.bonitasoft.engine.bpm.document.Document;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class ReadSheetConnectorTest {

    @Mock
    private Document mockDocument;

    private ReadSheetConnector connector;

    @BeforeEach
    void setUp() {
        connector = new ReadSheetConnector();
    }

    @Test
    void shouldFailValidationWhenDocumentMissing() {
        Map<String, Object> inputs = new HashMap<>();
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("document is mandatory");
    }

    @Test
    void shouldPassValidationWithDocument() throws Exception {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        connector.setInputParameters(inputs);

        assertThatCode(() -> connector.validateInputParameters())
                .doesNotThrowAnyException();
    }

    @Test
    void shouldBuildConfigurationWithDefaults() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        connector.setInputParameters(inputs);

        ExcelConfiguration config = connector.buildConfiguration();

        assertThat(config.getDocument()).isEqualTo(mockDocument);
        assertThat(config.getSheetIndex()).isEqualTo(0);
        assertThat(config.isEvaluateFormulas()).isTrue();
        assertThat(config.isIncludeHeaders()).isTrue();
    }

    @Test
    void shouldBuildConfigurationWithCustomValues() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("sheetName", "Data");
        inputs.put("sheetIndex", 2);
        inputs.put("evaluateFormulas", false);
        inputs.put("includeHeaders", false);
        connector.setInputParameters(inputs);

        ExcelConfiguration config = connector.buildConfiguration();

        assertThat(config.getSheetName()).isEqualTo("Data");
        assertThat(config.getSheetIndex()).isEqualTo(2);
        assertThat(config.isEvaluateFormulas()).isFalse();
        assertThat(config.isIncludeHeaders()).isFalse();
    }
}
