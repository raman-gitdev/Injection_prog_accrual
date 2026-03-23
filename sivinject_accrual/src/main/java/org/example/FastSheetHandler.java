package org.example;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

/**
 * SAX handler for XLSX sheets.
 * Fixed:
 * - cellValue now uses StringBuilder (was String +=, created garbage per characters() call)
 * - resolveValue is null-safe
 * - column bounds checked before put
 */
public class FastSheetHandler extends DefaultHandler {

    private final SharedStringsTable sst;
    private final StylesTable styles;
    private final Table table;

    private boolean isHeader = true;
    private int currentCol = -1;

    // FIX: was String cellValue with +=, which allocates a new String on every characters() call
    private final StringBuilder cellValue = new StringBuilder();

    private Map<String, Object> rowData;
    private int styleIndex = -1;
    private String cellType;

    FastSheetHandler(SharedStringsTable sst, StylesTable styles, Table table) {
        this.sst = sst;
        this.styles = styles;
        this.table = table;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {

        if ("row".equals(name)) {
            currentCol = -1;
            if (!isHeader) {
                rowData = table.newRow();
            }
        } else if ("c".equals(name)) {
            String cellRef = attributes.getValue("r");
            currentCol = columnIndexFromCellRef(cellRef);

            cellType = attributes.getValue("t");
            String styleIndexStr = attributes.getValue("s");
            styleIndex = (styleIndexStr != null) ? Integer.parseInt(styleIndexStr) : -1;

            cellValue.setLength(0); // reset for this cell

            if (isHeader) {
                while (table.getColumns().size() <= currentCol) {
                    table.addColumn("data" + (table.getColumns().size() + 1));
                }
            }
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        cellValue.append(ch, start, length); // no String allocation
    }

    @Override
    public void endElement(String uri, String localName, String name)
    {

        if ("v".equals(name)) {
            if (!isHeader && rowData != null) {
                if (currentCol >= 0 && currentCol < table.getColumns().size()) {
                    rowData.put(
                            table.getColumns().get(currentCol),
                            resolveValue(cellValue.toString(), cellType)
                    );
                }
            }
        }

        if ("row".equals(name)) {
            if (isHeader) {
                isHeader = false;
            } else {
                if (rowData != null) {
                    table.addRow(rowData);
                }
            }
        }
    }

    private int columnIndexFromCellRef(String cellRef)
    {
        if (cellRef == null || cellRef.isEmpty()) return 0;
        int col = 0;
        for (int i = 0; i < cellRef.length(); i++) {
            char c = cellRef.charAt(i);
            if (Character.isDigit(c)) break;
            col = col * 26 + (c - 'A' + 1);
        }
        return col - 1;
    }

    private boolean isNumeric(String str)
    {
        try
        {
            Double.parseDouble(str);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private Object resolveValue(String value, String cellType)
    {
        if (value == null || value.isEmpty())
            return null;value = value.trim();

        try {
            // Shared string
            if ("s".equals(cellType)) {
                int idx = Integer.parseInt(value);
                if (idx >= 0 && idx < sst.getCount()) {
                    return sst.getItemAt(idx).getString();
                }
                return null;
            }

            // Boolean
            if ("b".equals(cellType)) {
                return "1".equals(value);
            }

            // Error cell
            if ("e".equals(cellType)) {
                return value;
            }

            // Non-numeric inline string
            if (!isNumeric(value)) {
                return value;
            }

            // Numeric — check for date format
            if (styleIndex >= 0 && styles != null)
            {
                XSSFCellStyle style = styles.getStyleAt(styleIndex);
                if (style != null) {
                    short formatIndex = style.getDataFormat();
                    String formatString = style.getDataFormatString();
                    if (DateUtil.isADateFormat(formatIndex, formatString)) {
                        double numericValue = Double.parseDouble(value);
                        Date date = DateUtil.getJavaDate(numericValue);
                        return new SimpleDateFormat("yyyy-MM-dd").format(date);
                    }
                }
            }

            return value;

        } catch (Exception ex) {
            return value; // return raw on any parse failure
        }
    }
}