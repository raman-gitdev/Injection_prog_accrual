package org.example;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.channels.ScatteringByteChannel;
import java.sql.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class Main {

    static String glc_SIVFILEPath = "E:\\PRODUCTION\\DATA\\Accrual\\SIVFILE\\";


    // ── per-file state (reset at top of startproces) ──────────────────────────
    static int SheetCount = 1;
    static int currentSheetIndex = 0;
    static String CurrentSheetname = "";
    static List<String> sheetList = new ArrayList<>();
    static Map<String, Integer> physicalSheetIndex = new HashMap<>();
    static String glc_mapped_to = "";

    // ─────────────────────────────────────────────────────────────────────────

    public static void main(String[] args) {
        MainFrame frame = new MainFrame();
        frame.setVisible(true);

        List<String> dbname = new ArrayList<>();
        dbname.add("EM6_Accrual");

        for (String db : dbname) {
            doWork(db, frame);
        }


        System.exit(0);
    }

    private static void doWork(String dbsname, MainFrame frame)
    {
        Db.use(dbsname);

        List<Map<String, Object>> table = new ArrayList<>();
//
//        String sql = "SELECT a.*,FileType,'',b.frequency,b.sheet_name as ProcessStatus FROM sivfile a  left join carrmapping b on a.mapped_to = b.Carrier_Map_Name " +
//                "where status = 'MAPPED'  and sent_date >= '2025-09-01 14:25:12.000' and carrier not in ('corrections','match') and file_qtr='FY26Q2' " +
//                "and filetype='PCI' order by 1 desc"; //and a.id in (8209)
//
        String sql= "SELECT a.*,FileType,'',b.frequency,b.sheet_name as ProcessStatus FROM sivfile a  left join carrmapping b on a.mapped_to = b.Carrier_Map_Name " +
                "where a.id in (8175,8179,8182,8187,8204,8205,8206,8208,8209,8211,8223,8227,8228,8230,8232,8233,8234,8235,8236,8246,8254,8255,8257,8258,8259,8260,8271)";
                //PCI"where a.id in (8193,8194,8195,8196,8197,8198,8214,8215,8216,8217,8238,8239,8241,8199,8200,8203,8213,8218,8219,8220,8225,8242,8243,8240,8244,8245,8251)";//,8273,8262,8272
        //8268,8262,8263,8264,8266,8267,8265,8269,8273,8272,8270

        try (Connection conn = Db.getConnection()) {

            conn.createStatement().execute("SET search_path TO \"EM6_Accrual\"");

            try (PreparedStatement ps = conn.prepareStatement(sql);
                 ResultSet res = ps.executeQuery()) {

                ResultSetMetaData meta = res.getMetaData();
                int columnCount = meta.getColumnCount();
                while (res.next()) {
                    Map<String, Object> row = new HashMap<>();
                    for (int i = 1; i <= columnCount; i++) {
                        row.put(meta.getColumnName(i),res.getObject(i));
                    }
                    table.add(row);
                }

                frame.loadData(table);

                for (int i = 0; i < table.size(); i++)
                {

                    SivFileInfo fileInfo= new SivFileInfo();



                    Map<String, Object> row = table.get(i);
                    int fileid = ((Number) row.get("id")).intValue();

                    fileInfo.setFileName(row.get("original_name").toString());
                    fileInfo.setMappedTo(row.get("mapped_to").toString());
                    fileInfo.setFileType(row.get("filetype") == null ? "" :row.get("filetype").toString());
                    fileInfo.setFrequency(row.get("frequency") == null ? "" : row.get("frequency").toString());
                    fileInfo.setSheetName(row.get("sheet_name") == null ? "" :row.get("sheet_name").toString());

                    Object uid = row.get("uid_master");
                    fileInfo.setUidMaster(uid == null ? 0 : (int) row.get("uid_master"));
                    frame.updateStatus(i,"In process...");

                    ProcessLogger logger = new ProcessLogger();

                    boolean success= startproces(fileid,fileInfo,logger);

                    // Store full log so popup can show it
                    frame.setRowLog(i, logger.getFullLog());
                    if (!success) {
                        frame.updateStatus(i, "Failed — " + logger.getFailureMessage());
                        // Write error to DB
                        updateProcessStatus(fileid, "FAILED: " + logger.getFailureMessage());
                    } else {
                        frame.updateStatus(i, "Done");
                    }
                }

            }
        }
        catch (SQLException ex)
        {
            ex.printStackTrace();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // updateProcessStatus — writes error/success info to sivfile.processupdate
    // ─────────────────────────────────────────────────────────────────────────
    private static void updateProcessStatus(int fileId, String message)
    {
        String sql = "UPDATE sivfile SET UploadRemark = ?,uploaded_date=? WHERE id = ?";
        try (Connection conn = Db.getConnection()) {
            conn.createStatement().execute("SET search_path TO \"EM6_Accrual\"");
            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                // Truncate to 1000 chars to fit DB column safely
                String safeMsg = message.length() > 1000 ? message.substring(0, 1000) : message;
                ps.setString(1, safeMsg);
                ps.setTimestamp(2, java.sql.Timestamp.valueOf(LocalDateTime.now()));
                ps.setInt(3, fileId);
                ps.executeUpdate();
            }
        } catch (Exception ex) {
            System.err.println("Failed to update processupdate for file " + fileId + ": " + ex.getMessage());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // startproces
    // ─────────────────────────────────────────────────────────────────────────
    private static boolean startproces(int fileid,SivFileInfo info,ProcessLogger logger)
    {
        // ── reset all per-file static state ──────────────────────────────────
        CurrentSheetname = "";
        currentSheetIndex = 0;
        SheetCount = 1;
        sheetList.clear();
        physicalSheetIndex.clear();
        glc_mapped_to = "";

        String glc_numericColumns;
        String glc_DateColumns;
        String glc_DateFormat;
        String glc_carrier = null;
        String glc_Currency = null;
        int gln_sivFileID = fileid;
        String glc_FileType;
        String PCI_Mapped_To;
        String glc_Frequency;


        if (info == null) {
            logger.log("ERROR: SivFileInfo is null, aborting.");
            return false;
        }
        Map<String, List<Table>> dataSet = new HashMap<>();

        try
        {
            for (int SC = 0; SC < SheetCount; SC++) {

                logger.log("Downloading: " + info.getFileName());
                boolean downloadsuccess = DownloadFile(info.getFileName(), logger);

                if (!downloadsuccess) {
                    return false;
                }
                logger.log("Download complete");

                String glc_file_name = info.getFileName();
                glc_mapped_to = info.getMappedTo();
                glc_FileType = info.getFileType();
                glc_Frequency = info.getFrequency();
                PCI_Mapped_To = info.getMappedTo();

                if (glc_mapped_to != null) {
                    String[] parts = glc_mapped_to.split("_");
                    if (parts.length > 1) glc_carrier = parts[1];
                } else {
                    logger.log("ERROR: mapped_to is null");
                    return false;
                }
                // File extension
                String extension = glc_file_name.substring(glc_file_name.lastIndexOf(".")).toUpperCase();

                // ── STEP 2: Parse ─────────────────────────────────────────────
                logger.log("📄 Parsing sheet [" + (currentSheetIndex + 1) + "] — " + extension);

                Table CSVTable;
                try {
                    if (".CSV".equals(extension)) {
                        CurrentSheetname = "CSV";
                        CSVTable = getCSVDataTable(glc_SIVFILEPath + glc_file_name.trim());
                    } else if (".XLSB".equals(extension)) {
                        CSVTable = processXLSBSheet(glc_SIVFILEPath + glc_file_name.trim(), logger);
                    } else {
                        CSVTable = getXLSDataTable(glc_SIVFILEPath + glc_file_name.trim());
                    }
                } catch (Exception ex) {
                    logger.error("Parse failed", ex);
                    return false;
                }

                logger.log("✔ Parsed " + CSVTable.rowCount() + " rows, sheet: " + CurrentSheetname);

                // Trim columns > 250
                while (CSVTable.getColumns().size() > 250) {
                    CSVTable.removeColumn(250);
                }

                if ("MAP_TRANSPAK_SG".equalsIgnoreCase(glc_mapped_to)) {
                    glc_Currency = CSVTable.getRows().get(0).get(CSVTable.getColumns().get(10)).toString();
                }

                // ── STEP 3: Mapping ───────────────────────────────────────────
                logger.log("🗺 Applying mapping: " + glc_mapped_to);

                int headerToDelete = getHeaderCount(glc_mapped_to, glc_file_name);
                //if (".XLSB".equals(extension)) headerToDelete--;

                // Header adjustments (direct conversion)
                switch (glc_mapped_to) {
                    case "MAP_LANDSBERG":
                    case "MAP_LANDSBERG_NEW":
                    case "MAP_RISESUN_NEW":
                        headerToDelete -= 2;
                        break;
                    case "MAP_UPSSCSUS_NEW":
                        headerToDelete -= 6;
                        break;
                    case "MAP_DHLEXPRESS":
                        headerToDelete += 15;
                        break;
                    case "MAP_GXO":
                        headerToDelete += 3;
                        break;
                    case "MAP_ALBA_WEEL":
                        headerToDelete += 1;
                        break;
                    default:
                        headerToDelete = headerToDelete -1;
                }

                for (int i = 0; i < headerToDelete; i++) {
                    if (!CSVTable.getRows().isEmpty()) CSVTable.getRows().remove(0);
                }

                // Mapping table
                Table mapTable;
                if ("MAP_CEVA".equals(glc_mapped_to)) {
                    mapTable = getMappingTable(glc_mapped_to + "_" + CurrentSheetname);
                } else {
                    mapTable = getMappingTable(glc_mapped_to);
                }

                if (mapTable.getRows().isEmpty()) {
                    logger.log("ERROR: No mapping found for " + glc_mapped_to);
                    return false;
                }


                Map<String, Object> mapRow = mapTable.getRows().get(0);
                glc_numericColumns = mapRow.get("numposition").toString();
                glc_DateColumns = mapRow.get("dateposition").toString();
                glc_DateFormat = mapRow.get("dateformat").toString();

                if (glc_numericColumns.isEmpty() || glc_DateColumns.isEmpty() || glc_DateFormat.isEmpty()) {
                    logger.log("ERROR: Mapping config incomplete (numposition/dateposition/dateformat)");
                    return false;
                }

                // Rename CSV headers
                for (int i = 0; i < CSVTable.getColumns().size(); i++) {
                    String oldName = CSVTable.getColumns().get(i).toString().trim();
                    Object raw = mapRow.get(oldName);
                    String mappedName = (raw != null) ? raw.toString().trim() : "";

                    if (mappedName.isEmpty() || mappedName.equalsIgnoreCase("null")) {
                        mappedName = oldName.isEmpty() ? "" : "NS";
                    }
                    final String finalMapped = mappedName.toLowerCase();
                    for (Map<String, Object> r : CSVTable.getRows()) {
                        if (r.containsKey(oldName)) {
                            Object val = r.remove(oldName);
                            r.put(finalMapped, val);
                        }
                    }
                    CSVTable.getColumns().set(i, finalMapped);
                }

                // Remove NS columns
                for (int i = 0; i < CSVTable.getColumns().size(); i++) {
                    String col = CSVTable.getColumns().get(i);
                    if ("NS".equalsIgnoreCase(col) || col.trim().isEmpty()) {
                        CSVTable.removeColumn(i);
                        i--;
                    }
                }


                BigDecimal Billedamount = BigDecimal.ZERO;

                List<String> columns = CSVTable.getColumns();
                List<Map<String, Object>> rows = CSVTable.getRows();
                BigDecimal F1 = BigDecimal.ZERO;

                for (int iRow = 0; iRow < rows.size(); iRow++) {
                    Map<String, Object> row = rows.get(iRow);
                    String colName = "billed amount";
                    Object objValue = row.get(colName);
                    String value = objValue == null ? "0" : objValue.toString().trim();
                    BigDecimal F2 = new BigDecimal(value);

                    F1 = F1.add(F2);
                }

                Billedamount = F1;


                CleanDecimalFields(CSVTable, glc_numericColumns);

                CleanDateFields(CSVTable, glc_DateFormat, glc_DateColumns);


                logger.log("✔ Mapping applied — " + CSVTable.getColumns().size() + " columns remaining");

                logger.log("🔨 Building transaction rows");


                // Remove NS columns
//                for (int i = 0; i < CSVTable.getColumns().size(); i++) {
//                    String col = CSVTable.getColumns().get(i);
//                    if ("NS".equalsIgnoreCase(col) || col.trim().isEmpty()) {
//                        CSVTable.removeColumn(i);
//                        i--;
//                    }
//                }

                // ===============================
                // FINAL TRANSACTION BUILD SECTION
                // ===============================

                Table dt_Transaction = new Table();
                LocalDateTime addedDT = LocalDateTime.now();
                Set<String> carrierList = new HashSet<>();

                Map<String, Object> crowlname=CSVTable.newRow();

                String ccollname="";
                String cvallname="";

                try
                {

                for (Map<String, Object> csvRow : CSVTable.getRows()) {
                    Map<String, Object> txRow = dt_Transaction.newRow();
                    txRow.put("file_id", gln_sivFileID);
                    txRow.put("added_date", addedDT);

                    // ---------------------------
                    // NON-PCI LOGIC
                    // ---------------------------
                    if (!"PCI".equalsIgnoreCase(glc_FileType)) {

                        if (!"MAP_RISESUN".equalsIgnoreCase(glc_mapped_to)) {
                            txRow.put("carrier_name", glc_carrier);
                            txRow.put("carrier_grouping", glc_carrier);
                            txRow.put("frequency", glc_Frequency);
                            txRow.put("sheet_name", CurrentSheetname);
                            carrierList.add(glc_carrier);
                            if ("MAP_TRANSPAK_SG".equalsIgnoreCase(glc_mapped_to)) {
                                txRow.put("currency", glc_Currency);
                            }

                        } else {
                            txRow.put("carrier_grouping", csvRow.get("carrier_name"));
                            txRow.put("frequency", glc_Frequency);
                            txRow.put("sheet_name", CurrentSheetname);
                            carrierList.add(glc_carrier);
                        }

                    } else {
                        if (glc_mapped_to.contains("NV")) {
                            carrierList.add("enV_1");
                            txRow.put("mapped_to", PCI_Mapped_To);
                        } else if (glc_mapped_to.contains("EM6")) {
                            carrierList.add("em6_1");
                            txRow.put("mapped_to", PCI_Mapped_To);
                            SheetCount = 2;
                        } else {
                            carrierList.add(String.valueOf(csvRow.get("carrier code")));
                        }
                    }
                    if (CSVTable.hasColumn("ccv1")) {
                        txRow.put("description", csvRow.get("ccv1"));
                    }

                    // ---------------------------
                    // COLUMN MAPPING
                    // ---------------------------
                    for (String col : CSVTable.getColumns()) {

                        crowlname = csvRow;
                        ccollname =col;

                        Object rawVal = csvRow.get(col);
                        String strVal = rawVal == null ? "" : rawVal.toString().trim();

                        cvallname=strVal;

                        // INTEGER
                        if (isIntegerColumn(col)) {

                            if (strVal.isEmpty()) {
                                txRow.put(col, "WT_FLAG".toLowerCase().equalsIgnoreCase(col) ? 1 : 0);
                            } else {
                                txRow.put(col,
                                        Integer.parseInt(strVal.toLowerCase().replaceAll("[^0-9-]", ""))
                                );
                            }
                        }

                        // DECIMAL
                        else if (isDecimalColumn(col)) {
                            BigDecimal val = BigDecimal.ZERO;

                            cvallname=val.toString();

                            if (strVal.isEmpty()
                                    || strVal.equalsIgnoreCase("N/A")
                                    || strVal.toLowerCase().contains("balance")) {
//                                txRow.put(col, 0.0);
                                val = BigDecimal.ZERO;
                            } else {

                                String original = strVal.trim();
                                boolean isNegativeByParenthesis =
                                        (original.startsWith("(") && original.endsWith(")")) || original.startsWith("-");

                                String cleaned = strVal.toLowerCase()
                                        .replace("nt$", "")
                                        .replace("$", "")
                                        .replace(",", "")
                                        .replace("(", "")
                                        .replace(")", "")
                                        .replaceAll("[^0-9eE+\\\\-\\\\.]", "");

                                if (!cleaned.isEmpty() && !cleaned.equals("-")) {
                                    try {
                                        val = new BigDecimal(cleaned);
                                        if (isNegativeByParenthesis) {
                                            val = val.negate();
                                        }

                                    } catch (NumberFormatException nfe) {
                                        val = BigDecimal.ZERO;
                                    }
                                }
                            }

                            // Apply scale only if necessary
                            if ("cust_exng_rate".equalsIgnoreCase(col)) {
                                if (val.scale() > 6) {
                                    val = val.setScale(6, RoundingMode.HALF_UP);
                                }
                            } else {
                                if (val.scale() > 2) {
                                    val = val.setScale(2, RoundingMode.HALF_UP);
                                }
                            }

                            txRow.put(col.toLowerCase(), val);

                        }

                        // DATE
                        else if (isDateColumn(col)) {

                            LocalDateTime dt = null;

                            if (!strVal.isEmpty()
                                    && !strVal.equalsIgnoreCase("NULL")
                                    && !strVal.equals("1900-01-00")
                                    && !strVal.equals("44156")
                                    && !strVal.toLowerCase().contains("tba")
                                    && !strVal.toLowerCase().contains("total")
                                    && !strVal.toLowerCase().contains("invoice")) {

                                if (strVal.matches("\\d{4}-\\d{2}-\\d{2}")) {
                                    dt = LocalDate.parse(strVal).atStartOfDay();
                                } else if (strVal.matches(
                                        "\\d{4}-\\d{2}-\\d{2}([ T]\\d{2}:\\d{2}(:\\d{2})?)?")) {
                                    if (strVal.contains("T")) {
                                        dt = LocalDateTime.parse(strVal);
                                    } else {
                                        DateTimeFormatter fmt = strVal.length() == 16
                                                ? DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm")
                                                : DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                                        dt = LocalDateTime.parse(strVal, fmt);
                                    }
                                }
                            }

                            if (dt == null) dt = LocalDateTime.of(1900, 1, 1, 0, 0);
                            txRow.put(col.toLowerCase(), dt);

                        } else {
                            txRow.put(col, strVal);
                        }

                        if (col.equalsIgnoreCase("dim weight")) {
                            BigDecimal dimWeight = BigDecimal.ZERO;

                            if (strVal == "") {
                                strVal = "0";
                            }

                        }
                    }

                    // DIM WEIGHT fallback
                    if (txRow.containsKey("dim weight") || CSVTable.hasColumn("weight")) {
                        Object dimWt = txRow.get("dim weight");
                        BigDecimal dimWeight = BigDecimal.ZERO;
                        try {
                            if (dimWt instanceof BigDecimal) {
                                dimWeight = (BigDecimal) dimWt;
                            } else if (dimWt != null) {
                                String str = dimWt.toString().trim();
                                if (!str.isEmpty() && !str.equalsIgnoreCase("N/A")
                                        && !str.equalsIgnoreCase("NULL") && !str.equals("-")) {
                                    String cleaned = str.replaceAll("[^0-9eE+\\-.]", "");
                                    if (!cleaned.isEmpty() && !cleaned.equals("-")
                                            && !cleaned.equals(".")
                                            && cleaned.matches("[-+]?\\d*(\\.\\d+)?([eE][-+]?\\d+)?")) {
                                        dimWeight = new BigDecimal(cleaned);
                                    }
                                }
                            }
                        } catch (Exception ex) {
                            System.err.println("DIM WEIGHT parse failed | value=[" + dimWt + "]");
                            dimWeight = BigDecimal.ZERO;
                        }
                        txRow.put("dim weight", dimWeight);
                    }

                    dt_Transaction.addRow(txRow);
                }
            }
            catch(Exception exx)
            {
                String ddd=ccollname;
                String dddd=cvallname;
                Map<String, Object> sssss=crowlname;
                System.out.println(exx.getMessage());
            }


                BigDecimal Billedamount1=BigDecimal.ZERO;

                List<Map<String, Object>> rows1 = dt_Transaction.getRows();
                BigDecimal F11= BigDecimal.ZERO;

                for (int iRow = 0; iRow < rows1.size(); iRow++)
                {
                    Map<String, Object> row = rows1.get(iRow);
                    String colName = "billed amount";
                    Object objValue = row.get(colName);
                    String value = objValue == null ? "0" : objValue.toString().trim();
                    BigDecimal F22 = new BigDecimal(value);

                    F11= F11.add(F22);
                }

                Billedamount1=F11;

                // Make sure columns are registered from first tx row
                if (!dt_Transaction.getRows().isEmpty()) {
                    Map<String, Object> firstTx = dt_Transaction.getRows().get(0);
                    for (String k : firstTx.keySet()) {
                        dt_Transaction.addColumn(k);
                    }
                }





                dataSet.computeIfAbsent("dt_Transaction", k -> new ArrayList<>()).add(dt_Transaction);
                logger.log("✔ " + dt_Transaction.rowCount() + " transaction rows built");
            }

        }
        catch (Exception e)
        {
            logger.error("Unexpected error in processing", e);
            return false;
        }

        // ── STEP 5: Insert to DB ──────────────────────────────────────────────
        if (SheetCount > 0 && !dataSet.isEmpty()) {
            logger.log("Inserting into database...");
            boolean ok = copyToSivData(dataSet, glc_mapped_to, info.getFileType(), info.getMappedTo(), gln_sivFileID, logger);
            if (!ok) return false;
            logger.log("Insert complete");
        }

        return true;

    }

    private static Table processXLSBSheet(String filePath, ProcessLogger logger)
    {
        Table tt = new Table();

        try (OPCPackage pkg = OPCPackage.open(filePath)) {  // FIX: was never closed

            XSSFBReader r = new XSSFBReader(pkg);
            XSSFBSharedStringsTable sst = new XSSFBSharedStringsTable(pkg);
            XSSFBStylesTable xssfbStylesTable = r.getXSSFBStylesTable();

            XSSFBReader.SheetIterator it = (XSSFBReader.SheetIterator) r.getSheetsData();

            if (currentSheetIndex == 0) {
                List<String> allSheets = new ArrayList<>();
                while (it.hasNext()) {
                    try (InputStream is = it.next()) {
                        allSheets.add(it.getSheetName());
                    }
                }
                for (int i = 0; i < allSheets.size(); i++) {
                    physicalSheetIndex.put(allSheets.get(i), i);
                }
                applyFixSheetLogic(allSheets);
                SheetCount = sheetList.size();
            }

            if (currentSheetIndex >= sheetList.size()) return tt;

            if ("MAP_EM6_INPROCESS".equalsIgnoreCase(glc_mapped_to)
                    || "MAP_EM6_PAYABLE".equalsIgnoreCase(glc_mapped_to)) {
                glc_mapped_to += (currentSheetIndex == 0) ? "_PARENT" : "_CHILD";
            }

            CurrentSheetname = sheetList.get(currentSheetIndex);
            int physicalIndex = physicalSheetIndex.get(CurrentSheetname);
            logger.log("📋 Reading sheet: " + CurrentSheetname + " (index " + physicalIndex + ")");

            it = (XSSFBReader.SheetIterator) r.getSheetsData();

            for (int i = 0; i <= physicalIndex; i++) {
                try (InputStream is = it.next()) {
                    if (i == physicalIndex) {
                        SheetHandler handler = new SheetHandler(tt);
                        XSSFBSheetHandler sheetHandler = new XSSFBSheetHandler(
                                is, xssfbStylesTable,
                                it.getXSSFBSheetComments(),
                                sst, handler,
                                new DataFormatter(), false);
                        sheetHandler.parse();
                    }
                }
            }

            currentSheetIndex++;

        } catch (Exception ex) {
            logger.error("XLSB parse error", ex);
        }

        return tt;
    }


    private static void applyFixSheetLogic(List<String> allSheets)
    {
        sheetList.clear();

        if ("MAP_RYDER".equals(glc_mapped_to)) {
            sheetList.add("SUMMARY");
        } else if ("MAP_ALBA_WEEL".equals(glc_mapped_to)) {
            sheetList.add("Invoice List");
        } else if ("MAP_BHS".equals(glc_mapped_to)) {
            sheetList.add("BHS SOA");
        } else if ("MAP_DHL_US".equals(glc_mapped_to)) {
            sheetList.add("Aging");
        } else if ("MAP_EXEL_SCLA".equals(glc_mapped_to)) {
            sheetList.add("Master Trackers");
        } else if ("MAP_EXEL_TRACY".equals(glc_mapped_to)) {
            sheetList.add("Summary Unpaid");
        } else if ("MAP_LEADER".equals(glc_mapped_to)) {
            sheetList.add("Process");
        } else if ("MAP_GXO".equals(glc_mapped_to)) {
            sheetList.add("Invoice Details");
        } else if ("MAP_UPSSCSUS_NEW".equals(glc_mapped_to)) {
            sheetList.add("Aging");
        } else if ("MAP_LANDSBERG_NEW".equals(glc_mapped_to)) {
            sheetList.add("Sheet1");
        } else if ("MAP_UPSSCSUS".equals(glc_mapped_to)) {
            sheetList.add("SCS 01-23-24");
        } else if ("MAP_NIPPON_JP".equals(glc_mapped_to)) {
            sheetList.add("JP");
        } else if ("MAP_DBS_NEW3".equals(glc_mapped_to)) {
            sheetList.add("SOA");
        } else if ("MAP_DBS_NEW4".equals(glc_mapped_to)) {
            sheetList.add("Sheet1");
        } else if ("MAP_DBS_NEW2".equalsIgnoreCase(glc_mapped_to)) {
            sheetList.add("Sheet1");
        } else if ("MAP_YUANFAN_NEW".equalsIgnoreCase(glc_mapped_to)) {
            sheetList.add("YF-template");
        } else if ("MAP_EM6_INPROCESS".equals(glc_mapped_to)
                || "MAP_EM6_PAYABLE".equals(glc_mapped_to)) {
            sheetList.add("PARENT");
            sheetList.add("CHILD ACCT_CODE DETAILS");
        } else if ("MAP_DHL_GF_NEW".equals(glc_mapped_to)) {
            sheetList.add("data");
        } else if ("MAP_DSV".equals(glc_mapped_to)) {
            sheetList.add("Sheet1");
        } else if ("MAP_APEX".equals(glc_mapped_to)) {
            sheetList.add("Sheet1");
        } else if ("MAP_DHLEXPRESS".equals(glc_mapped_to)) {
            sheetList.add("DSO Level 5_ Transaction");
        } else if ("MAP_DHL_APAC_NEW".equals(glc_mapped_to)) {
            sheetList.add("Sheet1");
        } else {
            // Default: take all valid sheets
            for (String s : allSheets) {
                if (!s.toLowerCase().contains("filterdatabase")
                        && !s.toLowerCase().contains("print_area")) {
                    sheetList.add(s);
                }
            }
        }

        SheetCount = sheetList.size();
        currentSheetIndex = 0;
    }

    public static Table getXLSDataTable(String xlsFilePath) {
        Table table = new Table();
        ZipSecureFile.setMinInflateRatio(0.001);

        try (OPCPackage pkg = OPCPackage.open(new File(xlsFilePath))) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            if (currentSheetIndex == 0) {
                List<String> allSheets = new ArrayList<>();
                XSSFReader.SheetIterator nameIterator =
                        (XSSFReader.SheetIterator) reader.getSheetsData();
                while (nameIterator.hasNext()) {
                    try (InputStream is = nameIterator.next()) {
                        allSheets.add(nameIterator.getSheetName());
                    }
                }
                for (int i = 0; i < allSheets.size(); i++) {
                    physicalSheetIndex.put(allSheets.get(i), i);
                }
                applyFixSheetLogic(allSheets);
            }

            if ("MAP_CEVA".equalsIgnoreCase(glc_mapped_to)) {
                sheetList.remove("Unbilled");
            }

            if (currentSheetIndex >= sheetList.size() || sheetList.isEmpty()) return table;

            CurrentSheetname = sheetList.get(currentSheetIndex);
            int physicalIndex = physicalSheetIndex.get(CurrentSheetname);

            XSSFReader.SheetIterator sheetIterator =
                    (XSSFReader.SheetIterator) reader.getSheetsData();

            int idx = 0;
            while (sheetIterator.hasNext()) {
                InputStream sheetStream = sheetIterator.next();

                if (idx++ != physicalIndex) {
                    sheetStream.close();
                    continue;
                }

                SAXParserFactory factory = SAXParserFactory.newInstance();
                factory.setNamespaceAware(true);
                SAXParser saxParser = factory.newSAXParser();
                XMLReader parser = saxParser.getXMLReader();

                FastSheetHandler handler = new FastSheetHandler(sst, styles, table);
                parser.setContentHandler(handler);
                parser.parse(new InputSource(sheetStream));
                sheetStream.close();
                break;
            }

            currentSheetIndex++;
            SheetCount = sheetList.size();

        } catch (Exception ex) {
            ex.printStackTrace();
        }

        return table;
    }


    private static boolean DownloadFile(String remoteFile, ProcessLogger logger)
    {
        String fileDownloadPath = glc_SIVFILEPath + remoteFile;
        File file = new File(fileDownloadPath);

        if (file.exists()) {
            logger.log("✔ File already exists locally, skipping download");
            return true;
        }

        FTPClient ftp = new FTPClient();
        try {
            ftp.setConnectTimeout(100_000);
            ftp.setDataTimeout(120_000); // FIX: was missing — prevents infinite hang on transfer
            ftp.connect("depository.em6worldwide.com");

            boolean login = ftp.login("bipintl", "goregaon");
            if (!login) throw new RuntimeException("FTP login failed");

            ftp.enterLocalPassiveMode();
            ftp.setFileType(FTP.BINARY_FILE_TYPE);

            String remotePath = "/ACCRUAL/sivdata/in/" + remoteFile;
            try (InputStream input = ftp.retrieveFileStream(remotePath);
                 FileOutputStream output = new FileOutputStream(fileDownloadPath)) {

                if (input == null)
                    throw new RuntimeException("FTP stream null — file may not exist: " + remotePath);

                byte[] buffer = new byte[8192];
                int bytesRead;
                while ((bytesRead = input.read(buffer)) != -1) {
                    output.write(buffer, 0, bytesRead);
                }
            }

            ftp.completePendingCommand();
            ftp.logout();
            return true;

        } catch (Exception e) {
            logger.error("Download failed: " + remoteFile, e);
            return false;
        } finally {
            try { ftp.disconnect(); } catch (Exception ignored) {}
        }
    }

    private static boolean copyToSivData(
            Map<String, List<Table>> dsTransaction,
            String mapna, String fileType, String PCI_Mapped_To,
            int gln_sivFileID, ProcessLogger logger) {

        String effectiveMap = mapna;
        Connection conn = null;

        try {
            conn = Db.getConnection();
            conn.setAutoCommit(false);

            for (Map.Entry<String, List<Table>> entry : dsTransaction.entrySet()) {
                List<Table> tables = entry.getValue();

                for (Table dt : tables) {
                    String tableName;

                    // Mapping rotation for multi-sheet PCI files
                    if ("MAP_EM6_INPROCESS_CHILD".equals(mapna)) {
                        mapna = "MAP_EM6_INPROCESS_PARENT";
                    } else if ("MAP_EM6_INPROCESS_PARENT".equals(mapna)) {
                        mapna = "MAP_EM6_INPROCESS_CHILD";
                    } else if ("MAP_EM6_PAYABLE_CHILD".equals(mapna)) {
                        mapna = "MAP_EM6_PAYABLE_PARENT";
                    } else if ("MAP_EM6_PAYABLE_PARENT".equals(mapna)) {
                        mapna = "MAP_EM6_PAYABLE_CHILD";
                    } else if (mapna.contains("KSI")) {
                        mapna = "MAP_KSI";
                    }

                    if ("PCI".equalsIgnoreCase(fileType)) {
                        if (mapna.contains("PARENT") || mapna.contains("BILL")
                                || mapna.contains("NV_INPROCESS")) {
                            tableName = "DD_PARENT";
                            effectiveMap = PCI_Mapped_To;
                        } else {
                            tableName = "DD_CHILD";
                            effectiveMap = PCI_Mapped_To;
                        }
                    } else {
                        tableName = "CA";
                    }

                    List<String> columns = dt.getColumns();
                    if (columns.isEmpty()) continue;

                    String colList = columns.stream()
                            .map(c -> "\"" + c.toLowerCase() + "\"")
                            .collect(Collectors.joining(","));
                    String placeholders = String.join(",",
                            Collections.nCopies(columns.size(), "?"));

                    String sql = "INSERT INTO " + tableName +
                            " (" + colList + ") VALUES (" + placeholders + ")";

                    conn.createStatement().execute("SET search_path TO \"EM6_Accrual\"");

                    try (PreparedStatement ps = conn.prepareStatement(sql)) {
                        for (Map<String, Object> row : dt.getRows()) {
                            int idx = 1;
                            for (String col : columns) {
                                ps.setObject(idx++, row.get(col.toLowerCase()));
                            }
                            ps.addBatch();
                        }
                        ps.executeBatch();
                        logger.log("✔ Inserted " + dt.rowCount() + " rows into " + tableName);
                    }
                }
            }

            // Update sivfile status = PROCESSED
            try (PreparedStatement ps = conn.prepareStatement(
                    "UPDATE sivfile SET status = ? WHERE id = ?")) {
                ps.setString(1, "PROCESSED");
                //ps.setString(2, "Processed successfully at " + LocalDateTime.now());
                ps.setInt(2, gln_sivFileID);
                ps.executeUpdate();
            }

            conn.commit();
            return true;

        } catch (Exception ex) {
            logger.error("DB insert failed", ex);

            if (conn != null) {
                try { conn.rollback(); } catch (Exception ignored) {}
            }
            return false;
        } finally {
            if (conn != null) {
                try { conn.close(); } catch (Exception ignored) {}
            }
        }
    }



    public static Table getCSVDataTable(String csvFilePath) {
        Table table = new Table();
        try {
            char separator = detectSeparator(csvFilePath);
            CSVParser parser = new CSVParserBuilder()
                    .withSeparator(separator).withIgnoreQuotations(false).build();

            try (CSVReader reader = new CSVReaderBuilder(new FileReader(csvFilePath))
                    .withCSVParser(parser).build()) {

                String[] row;
                int blankCounter = 0;

                row = reader.readNext();
                if (row == null) return table;

                int colCount = Math.min(row.length, 300);
                for (int i = 1; i <= colCount; i++) table.addColumn("data" + i);

                Map<String, Object> firstRow = table.newRow();
                for (int i = 0; i < colCount; i++) {
                    firstRow.put(table.getColumns().get(i), clean(row[i]));
                }
                table.addRow(firstRow);

                while ((row = reader.readNext()) != null) {
                    boolean rowHasData = false;
                    Map<String, Object> dataRow = table.newRow();

                    for (int i = 0; i < Math.min(row.length, colCount); i++) {
                        String val = clean(row[i]);
                        if (val != null && !val.isEmpty()) { rowHasData = true; blankCounter = 0; }
                        dataRow.put(table.getColumns().get(i), val);
                    }
                    table.addRow(dataRow);

                    if (!rowHasData) blankCounter++;
                    if (blankCounter == 10) {
                        for (int i = 0; i < 10; i++) {
                            table.getRows().remove(table.getRows().size() - 1);
                        }
                        break;
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return table;
    }

    private static char detectSeparator(String filePath) throws IOException {
        char[] candidates = {',', ';', '\t', ':'};
        int[] counts = new int[candidates.length];
        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            String line = br.readLine();
            if (line == null) return ',';
            for (int i = 0; i < candidates.length; i++) {
                for (char c : line.toCharArray()) {
                    if (c == candidates[i]) counts[i]++;
                }
            }
        }
        int max = 0; char sep = ',';
        for (int i = 0; i < counts.length; i++) {
            if (counts[i] > max) { max = counts[i]; sep = candidates[i]; }
        }
        return sep;
    }

    private static String clean(String v) {
        if (v == null) return null;
        return v.replace("\"", "").trim();
    }



    private static int getHeaderCount(String mapName, String flname) {
        int headerCount = 0;
        String sql =
                "SELECT LTRIM(RTRIM(REPLACE(REPLACE(indexing,'header',''),'BY_PASS',''))) AS header " +
                        "FROM sivmaps WHERE sivprog = ? AND type_col = 'MAPPING'";
        try (Connection conn = Db.getConnection()) {
            conn.createStatement().execute("SET search_path TO \"EM6_Accrual\"");
            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setString(1, mapName);
                try (ResultSet rs = ps.executeQuery()) {
                    if (rs.next()) {
                        String headerStr = rs.getString("header");
                        if (headerStr == null || headerStr.trim().isEmpty()) {
                            headerCount = 0;
                        } else if ("MAP_UPSSP".equals(mapName) && flname.contains("0390JE")) {
                            headerCount = 7;
                        } else if ("MAP_UPSSCSTW".equals(mapName) && flname.contains("CMF")) {
                            headerCount = 3;
                        } else if ("MAP_UPSSCSTW".equals(mapName) && !flname.contains("CMF")) {
                            headerCount = 7;
                        } else {
                            String num = headerStr.replaceAll("\\D+", "");
                            headerCount = num.isEmpty() ? 0 : Integer.parseInt(num);
                        }
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return headerCount;
    }


    private static Table getMappingTable(String mapName) {
        Table table = new Table();
        String sql = "SELECT * FROM sivmaps WHERE lower(sivprog) = ? AND type_col = 'MAPPING' AND isactive = 1";
        try (Connection conn = Db.getConnection()) {
            conn.createStatement().execute("SET search_path TO \"EM6_Accrual\"");
            try (PreparedStatement ps = conn.prepareStatement(sql)) {
                ps.setString(1, mapName.toLowerCase());
                try (ResultSet rs = ps.executeQuery()) {
                    int columnCount = rs.getMetaData().getColumnCount();
                    for (int i = 1; i <= columnCount; i++) {
                        table.addColumn(rs.getMetaData().getColumnName(i));
                    }
                    while (rs.next()) {
                        Map<String, Object> row = table.newRow();
                        for (int i = 1; i <= columnCount; i++) {
                            row.put(rs.getMetaData().getColumnName(i), rs.getObject(i));
                        }
                        table.addRow(row);
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return table;
    }


    private static boolean isIntegerColumn(String col) {
        return col.equalsIgnoreCase("WT_FLAG") || col.equalsIgnoreCase("qty")
                || col.equalsIgnoreCase("pieces") || col.equalsIgnoreCase("count")
                || col.equalsIgnoreCase("number of pieces")
                || col.equalsIgnoreCase("total_tax_paid") || col.equalsIgnoreCase("miles");
    }

    private static boolean isDecimalColumn(String col) {
        return col.equalsIgnoreCase("weight") || col.equalsIgnoreCase("amt_USD")
                || col.equalsIgnoreCase("rate") || col.equalsIgnoreCase("price")
                || col.equalsIgnoreCase("cust_exng_rate") || col.equalsIgnoreCase("ca_orig_amt")
                || col.equalsIgnoreCase("billed_amount") || col.equalsIgnoreCase("passed_amount")
                || col.equalsIgnoreCase("billed amount") || col.equalsIgnoreCase("passed amount")
                || col.equalsIgnoreCase("seq number") || col.equalsIgnoreCase("sub seq number")
                || col.equalsIgnoreCase("ovc amount");
    }

    private static boolean isDateColumn(String col) {
        return col.toLowerCase().contains("date") || col.equalsIgnoreCase("added_date")
                || col.equalsIgnoreCase("invoice_date") || col.equalsIgnoreCase("ship_date")
                || col.equalsIgnoreCase("delivered_date") || col.equalsIgnoreCase("pro_date")
                || col.equalsIgnoreCase("ctsi_received_date") || col.equalsIgnoreCase("entry_date")
                || col.equalsIgnoreCase("paid_date") || col.equalsIgnoreCase("file_reception_date");
    }

    private static double round(double value, int scale) {
        return BigDecimal.valueOf(value)
                .setScale(scale, RoundingMode.HALF_UP)
                .doubleValue();
    }



    private static void CleanDateFields(Table csvTable,String glc_DateFormat,String glc_DateColumns) {
        try {
            int YCount = glc_DateFormat.split("Y").length - 1;
            int MCount = glc_DateFormat.split("M").length - 1;
            int DCount = glc_DateFormat.split("D", -1).length - 1;

            int YPos = glc_DateFormat.indexOf("Y");
            int MPos = glc_DateFormat.indexOf("M");
            int DPos = glc_DateFormat.indexOf("D");

            String[] arrFields = glc_DateColumns.split(",");

            List<String> columns = csvTable.getColumns();
            List<Map<String, Object>> rows = csvTable.getRows();

            for (int iRow = 0; iRow < rows.size(); iRow++) {
                Map<String, Object> row = rows.get(iRow);

                for (int i = 0; i < columns.size(); i++) {

                    String colName = columns.get(i);
                    String value = row.get(colName) == null ? "" : row.get(colName).toString().trim();

                    if (!Arrays.asList(arrFields).contains(String.valueOf(i + 1)))
                        continue;

                    if (value.isEmpty() || value.toLowerCase().contains("tba") || value.contains("$") || value.toLowerCase().contains("days")) {
                        row.put(colName, "");
                        continue;
                    }

                    if (value.matches("\\d{4}-\\d{2}-\\d{2}")) {
                        continue;
                    }

                    if (value.matches("\\d{4}-\\d{2}-\\d{2}([ T]\\d{2}:\\d{2}(:\\d{2})?)?")) {
                        continue;
                    }

                    if (value.matches("(?i).*jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec.*")) {

                        LocalDate parsed = parseMonthDate(value);
                        row.put(colName, parsed.toString());
                        continue;
                    }


                    // Example: convert "MM/DD/YYYY" to "YYYY-MM-DD"
                    if (!value.isEmpty() && value.contains("/")) {
                        String[] parts = value.split("/");
                        if (parts.length == 3) {
                            String formatted = String.format("%04d-%02d-%02d",
                                    Integer.parseInt(parts[2].length() == 2 ? "20" + parts[2] : parts[2]),
                                    Integer.parseInt(parts[0]),
                                    Integer.parseInt(parts[1])
                            );
                            row.put(colName, formatted);
                        }
                    } else if (glc_DateFormat.equals("YYYYMMDD") && value.length() == 6) {
                        row.put(colName, "20" + value);
                    }

                    // Extract year, month, day according to positions
                    if (row.get(colName) != null) {
                        String val = row.get(colName).toString();
                        if (val.length() >= YPos + YCount && val.length() >= MPos + MCount && val.length() >= DPos + DCount) {
                            String lc_Year = val.substring(YPos, YPos + YCount);
                            if (lc_Year.length() == 2) lc_Year = "20" + lc_Year;

                            String lc_Month = val.substring(MPos, MPos + MCount);
                            String lc_Day = val.substring(DPos, DPos + DCount);

                            row.put(colName, lc_Year + "-" + lc_Month + "-" + lc_Day);
                        }
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            // sendErrorMail equivalent here
        }
    }

    private static LocalDate parseMonthDate(String value) {

        DateTimeFormatter[] fmts = new DateTimeFormatter[]{
                DateTimeFormatter.ofPattern("dd-MMM-yy", Locale.ENGLISH),
                DateTimeFormatter.ofPattern("dd-MMM-yyyy", Locale.ENGLISH),
                DateTimeFormatter.ofPattern("MMM dd yyyy", Locale.ENGLISH)
        };

        for (DateTimeFormatter f : fmts) {
            try {
                return LocalDate.parse(value, f);
            } catch (Exception ignored) {}
        }
        throw new IllegalArgumentException("Invalid month date: " + value);
    }


    private static void CleanDecimalFields(Table csvTable,String glc_numericColumns)
    {
        try {
            String[] arrFields = glc_numericColumns.split(",");
            List<String> columns = csvTable.getColumns();
            List<Map<String, Object>> rows = csvTable.getRows();

            for (int iRow = 0; iRow < rows.size(); iRow++)
            {
                Map<String, Object> row = rows.get(iRow);

                for (int i = 0; i < columns.size(); i++)
                {
                    String colName = columns.get(i);
                    Object objValue = row.get(colName);
                    String value = objValue == null ? "" : objValue.toString().trim();

                    if (!Arrays.asList(arrFields).contains(String.valueOf(i + 1)))
                        continue;

                    if (value.isEmpty()) {
                        row.put(colName, 0);
                        continue;
                    }

                    // Remove all non-numeric except "." and "-"
                    value = value.replaceAll("[^\\d.,-]", "").replace(",", "");

                    double num = 0;
                    try {
                        num = Double.parseDouble(value);
                    } catch (Exception e) {
                        num = 0;
                    }

                    // Round to 2 decimal places
                    BigDecimal bd = BigDecimal.valueOf(num).setScale(2, RoundingMode.HALF_UP);
                    row.put(colName, bd.doubleValue());
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            // sendErrorMail equivalent here
        }
    }



}