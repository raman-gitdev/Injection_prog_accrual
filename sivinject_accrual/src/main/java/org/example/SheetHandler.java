package org.example;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.HashMap;
import java.util.*;



/**
 * Streaming XLSB sheet handler.
 * - No StringBuilder (was the source of OutOfMemoryError)
 * - First row becomes column headers (data1, data2 ...)
 * - Subsequent rows added to Table
 */


public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    // 500 covers any realistic spreadsheet; raise if you ever see >500 columns
    private static final int MAX_COLS = 500;

    private final Table table;
    private final String[] currentRow = new String[MAX_COLS];
    private boolean headerProcessed = false;
    private int maxSeenColumn = 0;
    public SheetHandler(Table table) {
        this.table = table;
    }

    @Override
    public void startRow(int rowNum) {
        // Reset array slice that was actually used — no need to clear the whole 500 slots
        if (maxSeenColumn > 0) {
            Arrays.fill(currentRow, 0, maxSeenColumn, null);
        }
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        int colIndex = parseColIndex(cellReference);
        if (colIndex < MAX_COLS) {
            currentRow[colIndex] = (formattedValue == null) ? "" : formattedValue;
            if (colIndex + 1 > maxSeenColumn) {
                maxSeenColumn = colIndex + 1;
            }
        }
    }

    @Override
    public void endRow(int rowNum) {

        // Skip completely empty rows
        if (maxSeenColumn == 0) return;

        List<String> cols = table.getColumns();

        // -------- HEADER ROW (first non-empty row) --------
        if (!headerProcessed) {
            for (int i = 1; i <= maxSeenColumn; i++) {
                table.addColumn("data" + i);
            }
            // First row IS data in your pipeline — store it as a data row
            Map<String, Object> firstRow = table.newRow();
            cols = table.getColumns(); // re-fetch after addColumn calls
            for (int i = 0; i < cols.size(); i++) {
                firstRow.put(cols.get(i), currentRow[i] != null ? currentRow[i] : "");
            }
            table.addRow(firstRow);
            headerProcessed = true;
            return;
        }

        // -------- DATA ROWS --------
        Map<String, Object> row = table.newRow();
        for (int i = 0; i < cols.size(); i++) {
            row.put(cols.get(i), currentRow[i] != null ? currentRow[i] : "");
        }
        table.addRow(row);
    }

    /**
     * Parses zero-based column index from a cell reference (e.g. "A1"→0, "Z1"→25, "AA1"→26, "BC47"→54)
     * without allocating any objects. This replaces new CellReference(ref).getCol()
     * which was being called once per cell — ~10 million times on a 128-col × 78k-row sheet.
     */
    private static int parseColIndex(String cellRef) {
        int col = 0;
        for (int i = 0; i < cellRef.length(); i++) {
            char c = cellRef.charAt(i);
            if (c < 'A') break; // hit the numeric row part
            col = col * 26 + (c - 'A' + 1);
        }
        return col - 1; // convert to zero-based
    }

}
