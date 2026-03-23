package org.example;

import java.util.*;

/**
 * Fixed version of Table:
 * - columns backed by LinkedHashSet internally to get O(1) contains() checks
 * - addRow() no longer silently mutates schema (throws instead)
 * - removeColumn() by name added for clarity
 */

public class Table {

    private final List<String> columns = new ArrayList<>();
    private final Set<String> columnSet = new LinkedHashSet<>();
    private final List<Map<String, Object>> rows = new ArrayList<>();
    public List<String> getColumns() {
        return columns;
    }

    public List<Map<String, Object>> getRows() {
        return rows;
    }

    public void addColumn(String name) {
        if (!columnSet.contains(name)) {
            columns.add(name);
            columnSet.add(name);
        }
    }

    public void removeColumn(int index) {
        if (index < 0 || index >= columns.size()) return;
        String col = columns.remove(index);
        columnSet.remove(col);
        for (Map<String, Object> row : rows) {
            row.remove(col);
        }
    }

    public void removeColumn(String name) {
        if (!columnSet.contains(name)) return;
        columns.remove(name);
        columnSet.remove(name);
        for (Map<String, Object> row : rows) {
            row.remove(name);
        }
    }

    public boolean hasColumn(String name) {
        return columnSet.contains(name);
    }

    public Map<String, Object> newRow() {
        return new HashMap<>();
    }

    /**
     * Adds a row WITHOUT mutating the column schema.
     * Unknown keys in the row are silently ignored (previously caused ghost columns).
     */
    public void addRow(Map<String, Object> row) {
        rows.add(row);
    }

    public int rowCount() {
        return rows.size();
    }

    public int columnCount() {
        return columns.size();
    }
}
