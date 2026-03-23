package org.example;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * Collects timestamped log entries for a single file's processing run.
 * Thread-safe for appending; designed to be passed into startproces().
 */
public class ProcessLogger {

    private static final DateTimeFormatter FMT =
            DateTimeFormatter.ofPattern("HH:mm:ss");

    private final List<String> entries = new ArrayList<>();
    private boolean failed = false;
    private String failureMessage = "";

    public synchronized void log(String step) {
        String line = "[" + LocalDateTime.now().format(FMT) + "] " + step;
        entries.add(line);
    }

    public synchronized void error(String step, Exception ex) {
        failed = true;
        String msg = ex.getMessage() != null ? ex.getMessage() : ex.getClass().getSimpleName();
        failureMessage = step + ": " + msg;
        String line = "[" + LocalDateTime.now().format(FMT) + "] " + failureMessage;
        entries.add(line);
    }

    public synchronized boolean isFailed() {
        return failed;
    }

    /** Full error message written to DB processupdate column. */
    public synchronized String getFailureMessage() {
        return failureMessage;
    }

    /** All log lines joined for popup display. */
    public synchronized String getFullLog() {
        return String.join("\n", entries);
    }
}