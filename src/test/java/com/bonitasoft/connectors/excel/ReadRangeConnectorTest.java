package com.bonitasoft.connectors.excel;

import static org.assertj.core.api.Assertions.*;

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
class ReadRangeConnectorTest {

    @Mock
    private Document mockDocument;

    private ReadRangeConnector connector;

    @BeforeEach
    void setUp() {
        connector = new ReadRangeConnector();
    }

    @Test
    void shouldFailValidationWhenDocumentMissing() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("range", "A1:B10");
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("document is mandatory");
    }

    @Test
    void shouldFailValidationWhenRangeMissing() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("range is mandatory");
    }

    @Test
    void shouldPassValidationWithDocumentAndRange() throws Exception {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("range", "A1:D10");
        connector.setInputParameters(inputs);

        assertThatCode(() -> connector.validateInputParameters())
                .doesNotThrowAnyException();
    }

    @Test
    void shouldBuildConfigurationWithRange() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("range", "B2:E5");
        connector.setInputParameters(inputs);

        ExcelConfiguration config = connector.buildConfiguration();

        assertThat(config.getRange()).isEqualTo("B2:E5");
        assertThat(config.isEvaluateFormulas()).isTrue();
    }
}
