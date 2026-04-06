package com.bonitasoft.connectors.excel;

import static org.assertj.core.api.Assertions.*;

import org.bonitasoft.engine.bpm.document.Document;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.*;

@ExtendWith(MockitoExtension.class)
class WriteRangeConnectorTest {

    @Mock
    private Document mockDocument;

    private WriteRangeConnector connector;

    @BeforeEach
    void setUp() {
        connector = new WriteRangeConnector();
    }

    @Test
    void shouldFailValidationWhenDocumentMissing() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("range", "A1:B10");
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("A"));
        inputs.put("data", data);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("document is mandatory");
    }

    @Test
    void shouldFailValidationWhenRangeMissing() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("A"));
        inputs.put("data", data);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("range is mandatory");
    }

    @Test
    void shouldFailValidationWhenDataMissing() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("range", "A1:B2");
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("data is mandatory");
    }

    @Test
    void shouldPassValidationWithAllRequiredInputs() throws Exception {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("range", "A1:B2");
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("X", "Y"));
        inputs.put("data", data);
        connector.setInputParameters(inputs);

        assertThatCode(() -> connector.validateInputParameters())
                .doesNotThrowAnyException();
    }

    @Test
    void shouldBuildConfigurationWithAllInputs() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("range", "C3:E5");
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList(1, 2));
        inputs.put("data", data);
        inputs.put("outputFileName", "custom.xlsx");
        connector.setInputParameters(inputs);

        ExcelConfiguration config = connector.buildConfiguration();

        assertThat(config.getRange()).isEqualTo("C3:E5");
        assertThat(config.getData()).hasSize(1);
        assertThat(config.getOutputFileName()).isEqualTo("custom.xlsx");
    }
}
