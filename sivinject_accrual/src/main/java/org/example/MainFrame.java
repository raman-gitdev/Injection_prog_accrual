package org.example;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.List;
import java.util.Map;

/**
 * Refactored MainFrame:
 * - Wider window with more visible columns
 * - Status column color-coded (green/red/orange)
 * - Double-click any row to show full processing log popup
 * - loadData() is EDT-safe
 * - Per-row log stored so popup shows correct file's log
 */

public class MainFrame extends JFrame {

    DefaultTableModel tableModel;
    private JTable table;

    // Stores the full log per row index so popup can display it
    private String[] rowLogs;


    public MainFrame() {
        setTitle("File Processing Monitor");
        setSize(1000, 550);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLocationRelativeTo(null);
        setBackground(new Color(245, 245, 250));

        tableModel = new DefaultTableModel(
                new Object[]{"ID", "File Name", "Mapped To", "File Type", "Status"}, 0
        ) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // read-only grid
            }
        };

        table = new JTable(tableModel);
        table.setRowHeight(26);
        table.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        table.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 13));
        table.getTableHeader().setBackground(new Color(30, 80, 160));
        table.getTableHeader().setForeground(Color.WHITE);
        table.setSelectionBackground(new Color(180, 210, 255));
        table.setGridColor(new Color(220, 220, 230));
        table.setShowGrid(true);
        table.setIntercellSpacing(new Dimension(1, 1));

        // Column widths
        table.getColumnModel().getColumn(0).setPreferredWidth(50);   // ID
        table.getColumnModel().getColumn(1).setPreferredWidth(240);  // File Name
        table.getColumnModel().getColumn(2).setPreferredWidth(200);  // Mapped To
        table.getColumnModel().getColumn(3).setPreferredWidth(80);   // File Type
        table.getColumnModel().getColumn(4).setPreferredWidth(300);  // Status
        // Color-coded status column renderer
        table.getColumnModel().getColumn(4).setCellRenderer(new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(
                    JTable t, Object value, boolean isSelected,
                    boolean hasFocus, int row, int column) {

                super.getTableCellRendererComponent(t, value, isSelected, hasFocus, row, column);
                String val = value == null ? "" : value.toString();

                if (!isSelected) {
                    if (val.startsWith("✅")) {
                        setBackground(new Color(220, 255, 220));
                        setForeground(new Color(0, 120, 0));
                    } else if (val.startsWith("❌")) {
                        setBackground(new Color(255, 220, 220));
                        setForeground(new Color(180, 0, 0));
                    } else if (val.startsWith("⏳")) {
                        setBackground(new Color(255, 250, 210));
                        setForeground(new Color(140, 100, 0));
                    } else {
                        setBackground(Color.WHITE);
                        setForeground(Color.DARK_GRAY);
                    }
                }
                return this;
            }
        });

        // Double-click → show full log popup
        table.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                    int row = table.getSelectedRow();
                    if (row >= 0 && rowLogs != null && row < rowLogs.length) {
                        showLogPopup(row);
                    }
                }
            }
        });

        // Hint label
        JLabel hint = new JLabel("  💡 Double-click any row to view full processing log");
        hint.setFont(new Font("Segoe UI", Font.ITALIC, 12));
        hint.setForeground(new Color(90, 90, 120));
        hint.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setBorder(BorderFactory.createLineBorder(new Color(200, 200, 220)));

        add(scrollPane, BorderLayout.CENTER);
        add(hint, BorderLayout.SOUTH);
    }

    /**
     * Load all file rows into the grid. Must be called from any thread — EDT-safe.
     */

    public void loadData(List<Map<String, Object>> tableData) {
        rowLogs = new String[tableData.size()];

        SwingUtilities.invokeLater(() -> {
            tableModel.setRowCount(0);
            for (Map<String, Object> row : tableData) {
                tableModel.addRow(new Object[]{
                        row.get("id"),
                        row.get("original_name"),
                        row.get("mapped_to"),
                        row.get("filetype"),
                        ""   // status starts empty
                });
            }
        });
    }

    /**
     * Update the status cell for a given row. Call from any thread.
     * prefix: "⏳" for in-progress, "✅" for done, "❌" for failed.
     */
    public void updateStatus(int rowIndex, String status) {
        SwingUtilities.invokeLater(() -> {
            if (rowIndex < tableModel.getRowCount()) {
                tableModel.setValueAt(status, rowIndex, 4);
            }
        });
    }

    /**
     * Store the full processing log for a row so popup can show it.
     */
    public void setRowLog(int rowIndex, String log) {

        if (rowLogs == null) {
            rowLogs = new String[tableModel.getRowCount()];
        }
        if (rowIndex >= 0 && rowIndex < rowLogs.length) {
            rowLogs[rowIndex] = (log == null || log.isEmpty()) ? "(No steps logged)" : log;
        }
    }

    private void showLogPopup(int rowIndex) {
        String fileName  = tableModel.getValueAt(rowIndex, 1) + "";
        String mappedTo  = tableModel.getValueAt(rowIndex, 2) + "";
        String log       = rowLogs[rowIndex];

        JDialog dialog = new JDialog(this, "Processing Log — " + fileName, true);
        dialog.setSize(700, 500);
        dialog.setLocationRelativeTo(this);
        dialog.setLayout(new BorderLayout(8, 8));

        // Header
        JLabel header = new JLabel("  " + fileName + "  [" + mappedTo + "]");
        header.setFont(new Font("Segoe UI", Font.BOLD, 14));
        header.setForeground(new Color(30, 60, 130));
        header.setBorder(BorderFactory.createEmptyBorder(10, 10, 0, 10));

        // Log text area
        JTextArea logArea = new JTextArea(log == null ? "(No log recorded)" : log);
        logArea.setEditable(false);
        logArea.setFont(new Font("Consolas", Font.PLAIN, 12));
        logArea.setBackground(new Color(20, 20, 30));
        logArea.setForeground(new Color(180, 255, 180));
        logArea.setCaretColor(Color.WHITE);
        logArea.setMargin(new Insets(10, 12, 10, 12));

        JScrollPane scroll = new JScrollPane(logArea);
        scroll.setBorder(BorderFactory.createEmptyBorder(0, 10, 0, 10));

        // Close button
        JButton close = new JButton("Close");
        close.setFont(new Font("Segoe UI", Font.BOLD, 12));
        close.addActionListener(e -> dialog.dispose());
        JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btnPanel.add(close);

        dialog.add(header, BorderLayout.NORTH);
        dialog.add(scroll, BorderLayout.CENTER);
        dialog.add(btnPanel, BorderLayout.SOUTH);
        dialog.setVisible(true);
    }



}
