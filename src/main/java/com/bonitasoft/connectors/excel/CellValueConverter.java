package com.bonitasoft.connectors.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 * Bidirectional type mapping between Apache POI cells and Java objects.
 * Handles reading cell values to Java types and writing Java objects to cells.
 */
@Slf4j
public class CellValueConverter {

    private CellValueConverter() {
        // utility class
    }

    /**
     * Reads a cell value and converts it to an appropriate Java object.
     * Supports STRING, NUMERIC (including dates), BOOLEAN, FORMULA, BLANK, and ERROR types.
     *
     * @param cell the cell to read
     * @param evaluator formula evaluator (may be null to skip formula evaluation)
     * @return the Java object representation of the cell value
     */
    public static Object readCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return null;
        }

        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA && evaluator != null) {
            cellType = evaluator.evaluateFormulaCell(cell);
        }

        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue();
                }
                double numValue = cell.getNumericCellValue();
                // Return integer if the value has no fractional part
                if (numValue == Math.floor(numValue) && !Double.isInfinite(numValue)) {
                    long longVal = (long) numValue;
                    if (longVal >= Integer.MIN_VALUE && longVal <= Integer.MAX_VALUE) {
                        return (int) longVal;
                    }
                    return longVal;
                }
                return numValue;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
                return null;
            case ERROR:
                return "#ERROR(" + cell.getErrorCellValue() + ")";
            case FORMULA:
                // Formula not evaluated (evaluator was null)
                return cell.getCellFormula();
            default:
                return null;
        }
    }

    /**
     * Writes a Java object value into an Apache POI cell.
     *
     * @param cell the cell to write to
     * @param value the value to write
     */
    public static void writeCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof java.time.LocalDateTime) {
            cell.setCellValue((java.time.LocalDateTime) value);
        } else if (value instanceof java.time.LocalDate) {
            cell.setCellValue((java.time.LocalDate) value);
        } else if (value instanceof java.util.Date) {
            cell.setCellValue((java.util.Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }
}
