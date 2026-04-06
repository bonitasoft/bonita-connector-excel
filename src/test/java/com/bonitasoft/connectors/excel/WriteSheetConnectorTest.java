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
class WriteSheetConnectorTest {

    @Mock
    private Document mockDocument;

    private WriteSheetConnector connector;

    @BeforeEach
    void setUp() {
        connector = new WriteSheetConnector();
    }

    @Test
    void shouldFailValidationWhenDocumentMissing() {
        Map<String, Object> inputs = new HashMap<>();
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("A", "B"));
        inputs.put("data", data);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("document is mandatory");
    }

    @Test
    void shouldFailValidationWhenDataMissing() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("data is mandatory");
    }

    @Test
    void shouldFailValidationWhenDataEmpty() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        inputs.put("data", new ArrayList<>());
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class)
                .hasMessageContaining("data is mandatory");
    }

    @Test
    void shouldPassValidationWithDocumentAndData() throws Exception {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("Name", "Value"));
        inputs.put("data", data);
        connector.setInputParameters(inputs);

        assertThatCode(() -> connector.validateInputParameters())
                .doesNotThrowAnyException();
    }

    @Test
    void shouldBuildConfigurationWithDefaults() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("document", mockDocument);
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("A"));
        inputs.put("data", data);
        connector.setInputParameters(inputs);

        ExcelConfiguration config = connector.buildConfiguration();

        assertThat(config.getOutputFileName()).isEqualTo("output.xlsx");
        assertThat(config.isCreateSheetIfMissing()).isTrue();
        assertThat(config.getStreamingThreshold()).isEqualTo(10000);
    }
}
