package org.example;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.io.InputStream;
import java.util.Properties;

public class Db {

    private static final Properties props = new Properties();
    private static String currentDb;

    // STATIC BLOCK: runs once when class loads
    static {
        try (InputStream is =
                     Db.class.getClassLoader().getResourceAsStream("db.properties")) {

            if (is == null) {
                throw new RuntimeException("db.properties not found in resources");
            }

            props.load(is);

        } catch (Exception e) {
            throw new RuntimeException("Failed to load db.properties", e);
        }
    }

    // Called at app startup (like selecting connection string)
    public static void use(String dbName) {
        currentDb = dbName;
    }

    // Get JDBC connection for selected DB
    public static Connection getConnection() throws SQLException {

        if (currentDb == null) {
            throw new IllegalStateException("Database not selected. Call Db.use(dbName)");
        }

        String url = props.getProperty(currentDb + ".url");
        String urls = props.getProperty(currentDb+".host");
        String user = props.getProperty(currentDb + ".user");
        String password = props.getProperty(currentDb + ".password");

        if (url == null || user == null) {
            throw new RuntimeException("DB config not found for: " + currentDb);
        }


        return DriverManager.getConnection(url, user, password);
    }
}
