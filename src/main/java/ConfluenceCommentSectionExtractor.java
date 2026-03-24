import com.jakewharton.fliptables.FlipTable;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.FileOutputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Properties;

/**
 * Extracts Confluence page comments (inline + footer) and maps each to its
 * page section (the nearest preceding heading in the page body).
 *
 * <p>This utility reads an input Excel file containing rows of pages to process
 * (module metadata + Confluence pageId), retrieves the page body and comments
 * from Confluence via the REST API, maps inline comments to the nearest
 * preceding heading (section), and writes all comments into a single output
 * Excel workbook. It also prints a summary table to the console.</p>
 *
 * Dependencies (Maven / pom.xml provided separately):
 *   - org.apache.poi:poi-ooxml:5.2.5
 *   - org.json:json:20231013
 *   - org.jsoup:jsoup:1.17.2
 */
public class ConfluenceCommentSectionExtractor {

    // -----------------------------------------------------------------------
    // Configuration — update these before running
    // -----------------------------------------------------------------------
    private static String BASE_URL;
    private static String EMAIL;
    private static String API_TOKEN;
    
    private static String AUTH_HEADER;
    
    static {
        // Load configuration from config.properties at class initialization time.
        try (java.io.InputStream is = new java.io.FileInputStream("config.properties")) {
            Properties props = new Properties();
            props.load(is);
            BASE_URL  = props.getProperty("base.url");
            EMAIL     = props.getProperty("email");
            API_TOKEN = props.getProperty("api.token");
        } catch (java.io.IOException e) {
            throw new ExceptionInInitializerError("Failed to load config.properties: " + e.getMessage());
        }
    }

    // -----------------------------------------------------------------------
    // Entry point
    // -----------------------------------------------------------------------
    /**
     * Program main entry point. Expects a single argument: path to the input
     * Excel file. Reads the input rows, processes each page (fetching comments),
     * writes a consolidated comments Excel file and prints a summary to stdout.
     *
     * @param args program arguments. args[0] must be the input Excel file path.
     * @throws Exception on unexpected failures during processing or I/O
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            System.err.println("Usage: java ConfluenceCommentSectionExtractor <inputExcelFile>");
            System.exit(1);
        }
        String raw = EMAIL + ":" + API_TOKEN;
        AUTH_HEADER = "Basic " + Base64.getEncoder().encodeToString(raw.getBytes());

        String inputFile = args[0];
        List<InputRow> inputRows = readInputExcel(inputFile);

        List<CommentRecord> allComments = new ArrayList<>();
        List<SummaryRecord> summaryRows = new ArrayList<>();

        for (InputRow row : inputRows) {
            SummaryRecord sr = new SummaryRecord();
            try {
                System.out.println("Processing pageId: " + row.pageId +
                                 " (Module: " + row.module + ", Function: " + row.functionName + ")");
                List<CommentRecord> pageComments = processPage(row);
                allComments.addAll(pageComments);

                sr.module = row.module;
                sr.legacyId = row.legacyId;
                sr.newId = row.newId;
                sr.functionName = row.functionName;
                sr.functionType = row.functionType;
                sr.pageId = row.pageId;
                sr.numberOfComments = pageComments.size();
                summaryRows.add(sr);
            } catch (Exception e) {
                System.err.println("Error processing page " + row.pageId + ": " + e.getMessage());

                // Keep failed pages in summary with 0 comments
                sr.module = row.module;
                sr.legacyId = row.legacyId;
                sr.newId = row.newId;
                sr.functionName = row.functionName;
                sr.functionType = row.functionType;
                sr.pageId = row.pageId;
                sr.numberOfComments = 0;
                summaryRows.add(sr);
            }
        }

        // Write all comments to a single Excel file
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String outputPath = "d:/outputs/all_confluence_comments_" + timestamp + ".xlsx";
        writeExcel(allComments, outputPath);
        printExecutionSummary(summaryRows);
        printExecutionSummaryTabular(summaryRows);

        System.out.println();
        System.out.println("Done! Output: " + outputPath);
    }

    /**
     * Read the first sheet of the provided Excel workbook and transform each
     * non-empty row into an {@link InputRow} object. The method expects the
     * following column order: Module, Legacy ID, New ID, Function Name,
     * Function Type, pageId. The header row (row 0) is skipped.
     *
     * @param filePath absolute or relative path to the input .xlsx file
     * @return list of populated InputRow objects (skips rows with empty pageId)
     * @throws Exception on I/O or workbook parsing errors
     */
    private static List<InputRow> readInputExcel(String filePath) throws Exception {
        List<InputRow> rows = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Skip header row
                Row row = sheet.getRow(i);
                if (row == null) continue;

                InputRow inputRow = new InputRow();
                inputRow.module = getCellValue(row, 0);
                inputRow.legacyId = getCellValue(row, 1);
                inputRow.newId = getCellValue(row, 2);
                inputRow.functionName = getCellValue(row, 3);
                inputRow.functionType = getCellValue(row, 4);
                inputRow.pageId = getCellValue(row, 5);

                if (!inputRow.pageId.isEmpty()) {
                    rows.add(inputRow);
                }
            }
        }

        System.out.println("Loaded " + rows.size() + " page(s) from input Excel");
        return rows;
    }

    /**
     * Safely read a cell value as String. Supports STRING and NUMERIC types;
     * for numeric cells the value is converted to a long and then to String to
     * avoid decimal artifacts for integer page ids.
     *
     * @param row the worksheet row containing the cell
     * @param colIndex 0-based column index
     * @return the cell value as a non-null String (empty string when cell is
     *         missing or of unsupported type)
     */
    private static String getCellValue(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue());
            default:
                return "";
        }
    }

    /**
     * Process a single input row: fetch the page storage body, build a map of
     * inline marker -> section (nearest preceding heading), then fetch both
     * inline and footer comments and return them as a list of
     * {@link CommentRecord}.
     *
     * @param inputRow metadata and pageId to process
     * @return list of CommentRecord objects found on the page (may be empty)
     * @throws Exception on unexpected API errors or parsing problems
     */
    private static List<CommentRecord> processPage(InputRow inputRow) throws Exception {
        String pageId = inputRow.pageId;
        System.out.println("  Fetching page body (storage format)...");
        String storageBody = fetchPageBody(pageId);

        System.out.println("  Parsing headings and inline markers...");
        LinkedHashMap<String, List<String>> sectionMarkerMap = buildSectionMarkerMap(storageBody);
        Map<String, String> markerToSection = new HashMap<>();
        for (Map.Entry<String, List<String>> e : sectionMarkerMap.entrySet()) {
            for (String ref : e.getValue()) {
                markerToSection.put(ref, e.getKey());
            }
        }

        System.out.println("  Fetching comments...");
        List<CommentRecord> comments = new ArrayList<>();
        comments.addAll(fetchInlineComments(inputRow, markerToSection, pageId));
        comments.addAll(fetchFooterComments(inputRow, pageId));
        System.out.println("  Found " + comments.size() + " comment(s)");

        return comments;
    }

    /**
     * Print a compact CSV-like execution summary to stdout. This is primarily
     * useful for scripted consumption or simple logs; for a pretty table see
     * {@link #printExecutionSummaryTabular(List)}.
     *
     * @param rows list of per-page summary records to print
     */
    private static void printExecutionSummary(List<SummaryRecord> rows) {
        System.out.println();
        System.out.println("Module, Legacy ID, New ID, Function Name, "
                + "Function Type, pageId, Number of comments");
        for (SummaryRecord r : rows) {
            System.out.println(String.join(", ",
                    safe(r.module),
                    safe(r.legacyId),
                    safe(r.newId),
                    safe(r.functionName),
                    safe(r.functionType),
                    safe(r.pageId),
                    String.valueOf(r.numberOfComments)
            ));
        }
    }

    /**
     * Print a nicely formatted ASCII table of the execution summary using
     * the external FlipTable library. Headers and rows are derived from the
     * provided SummaryRecord list.
     *
     * @param rows list of per-page summary records to print in a tabular view
     */
    private static void printExecutionSummaryTabular(List<SummaryRecord> rows) {
        String[] headers = {
                "Module", "Legacy ID", "New ID", "Function Name",
                "Function Type", "pageId", "Number of comments"
        };

        String[][] data = new String[rows.size()][headers.length];
        for (int i = 0; i < rows.size(); i++) {
            SummaryRecord r = rows.get(i);
            data[i][0] = safe(r.module);
            data[i][1] = safe(r.legacyId);
            data[i][2] = safe(r.newId);
            data[i][3] = safe(r.functionName);
            data[i][4] = safe(r.functionType);
            data[i][5] = safe(r.pageId);
            data[i][6] = String.valueOf(r.numberOfComments);
        }

        System.out.println();
        System.out.println(FlipTable.of(headers, data));
    }

    /**
     * Null-safe helper that returns an empty string for null inputs.
     *
     * @param s input string
     * @return original string or empty string when null
     */
    private static String safe(String s) {
        return s == null ? "" : s;
    }


    // -----------------------------------------------------------------------
    // Step 1 — fetch page body in storage format
    // -----------------------------------------------------------------------
    /**
     * Fetch the Confluence page storage-format body for the given pageId.
     * This calls the Confluence REST API: /rest/api/content/{id}?expand=body.storage
     *
     * @param pageId Confluence page id
     * @return storage-format HTML/XML string of the page body
     * @throws Exception on HTTP or JSON parsing errors
     */
    private static String fetchPageBody(String pageId) throws Exception {
        String url = BASE_URL + "/rest/api/content/" + pageId
                + "?expand=body.storage";
        JSONObject json = getJson(url);
        return json.getJSONObject("body")
                .getJSONObject("storage")
                .getString("value");
    }

    // -----------------------------------------------------------------------
    // Step 2 — parse headings and inline markers in document order
    //
    // The storage format looks like:
    //   <h2>Section Title</h2>
    //   <p>Some text <ac:inline-comment-marker ac:ref="uuid1">highlighted</ac:inline-comment-marker></p>
    //   <h3>Sub-section</h3>
    //   ...
    //
    // We walk all elements in order; when we see a heading we start a new
    // section bucket; when we see a marker we add it to the current bucket.
    // -----------------------------------------------------------------------
    /**
     * Parse the storage-format page body and build a map of section title ->
     * list of inline comment marker references that appear under that section.
     * The iteration preserves document order by using a LinkedHashMap.
     *
     * @param storageXml page storage-format content (XML/HTML)
     * @return LinkedHashMap where keys are section titles and values are lists
     *         of inline comment marker references found in that section
     */
    static LinkedHashMap<String, List<String>> buildSectionMarkerMap(String storageXml) {
        // Jsoup handles XHTML/XML well enough for this purpose
        Document doc = Jsoup.parse(storageXml, "", org.jsoup.parser.Parser.xmlParser());

        LinkedHashMap<String, List<String>> map = new LinkedHashMap<>();
        String currentSection = "(Before first heading)";
        map.put(currentSection, new ArrayList<>());

        // Walk all elements in document order
        for (Element el : doc.getAllElements()) {
            String tag = el.tagName().toLowerCase();

            // Heading?
            if (tag.matches("h[1-6]")) {
                currentSection = el.text().trim();
                if (currentSection.isEmpty()) currentSection = "(Unnamed heading)";
                map.putIfAbsent(currentSection, new ArrayList<>());
                // Note: if two headings share the same text the second one will
                // reuse the same bucket — acceptable for typical pages.
                continue;
            }

            // Inline comment marker?
            if (tag.equals("ac:inline-comment-marker")) {
                String ref = el.attr("ac:ref");
                if (!ref.isEmpty()) {
                    map.get(currentSection).add(ref);
                }
            }
        }
        return map;
    }

    // -----------------------------------------------------------------------
    // Step 3a — fetch inline comments
    // -----------------------------------------------------------------------
    /**
     * Fetch inline comments for a page. Uses the Confluence comments API and
     * matches inline comments to their section using the provided
     * markerToSection map which maps inline marker references to section
     * titles.
     *
     * @param inputRow input metadata associated with the page
     * @param markerToSection mapping of inline marker reference -> section
     * @param pageId Confluence page id
     * @return list of CommentRecord instances for inline comments
     * @throws Exception on HTTP/JSON errors
     */
    private static List<CommentRecord> fetchInlineComments(
            InputRow inputRow, Map<String, String> markerToSection, String pageId) throws Exception {

        List<CommentRecord> results = new ArrayList<>();
        int start = 0;
        int limit = 50;

        while (true) {
            String url = BASE_URL + "/rest/api/content/" + pageId
                    + "/child/comment"
                    + "?expand=body.view,version,history,extensions.inlineProperties"
                    + "&depth=all"
                    + "&limit=" + limit
                    + "&start=" + start;

            JSONObject response = getJson(url);
            JSONArray results_ = response.getJSONArray("results");
            if (results_.isEmpty()) break;

            for (int i = 0; i < results_.length(); i++) {
                JSONObject c = results_.getJSONObject(i);

                JSONObject extensions = c.optJSONObject("extensions");
                if (extensions == null) continue;
                JSONObject inlineProps = extensions.optJSONObject("inlineProperties");
                if (inlineProps == null) continue;

                String markerRef        = inlineProps.optString("ref", "");
                String originalSelection = inlineProps.optString("originalSelection", "");

                String section = markerToSection.getOrDefault(markerRef, "(Section not resolved)");
                String id      = c.getString("id");
                String author  = "(Unknown)";
                String created = "(Unknown)";
                try {
                    author = c.getJSONObject("history")
                            .getJSONObject("createdBy")
                            .getString("displayName");
                } catch (Exception e) {
                    System.out.println(c.getJSONObject("history").toString());
                    System.err.println("    Warning: failed to get author for comment " + id + ": " + e.getMessage());
                }
                try {
                    created = c.getJSONObject("history")
                            .getString("createdDate");
                } catch (Exception e) {
                    System.out.println(c.getJSONObject("history").toString());
                    System.err.println("    Warning: failed to get createdDate for comment " + id + ": " + e.getMessage());
                }
                String updated = c.getJSONObject("version")
                        .getString("when");
                String bodyHtml = c.getJSONObject("body")
                        .getJSONObject("view")
                        .getString("value");
                String commentText = stripHtml(bodyHtml);

                String status = extensions.optString("resolution", "open");

                String link = BASE_URL + "/pages/viewpage.action?pageId=" + pageId
                        + "&focusedCommentId=" + id + "#comment-" + id;

                CommentRecord rec = new CommentRecord();
                // Populate input metadata
                rec.module = inputRow.module;
                rec.legacyId = inputRow.legacyId;
                rec.newId = inputRow.newId;
                rec.functionName = inputRow.functionName;
                rec.functionType = inputRow.functionType;

                // Populate comment fields
                rec.type              = "Inline";
                rec.section           = section;
                rec.originalSelection = originalSelection;
                rec.id                = id;
                rec.author            = author;
                rec.createdDate       = created;
                rec.lastUpdated       = updated;
                rec.commentText       = commentText;
                rec.status            = status;
                rec.link              = link;
                results.add(rec);
            }

            JSONObject links = response.optJSONObject("_links");
            if (links == null || !links.has("next")) break;
            start += limit;
        }
        return results;
    }

    // -----------------------------------------------------------------------
    // Step 3b — fetch footer (page-level) comments
    // -----------------------------------------------------------------------
    /**
     * Fetch top-level (footer) comments for a page. These comments do not have
     * inlineProperties and are considered page-level feedback. The method
     * returns a list of CommentRecord objects representing each footer comment.
     *
     * @param inputRow input metadata associated with the page
     * @param pageId Confluence page id
     * @return list of CommentRecord objects for footer comments
     * @throws Exception on HTTP/JSON errors
     */
    private static List<CommentRecord> fetchFooterComments(InputRow inputRow, String pageId) throws Exception {
        List<CommentRecord> results = new ArrayList<>();
        int start = 0;
        int limit = 50;

        while (true) {
            String url = BASE_URL + "/rest/api/content/" + pageId
                    + "/child/comment"
                    + "?expand=body.view,version,history,extensions.inlineProperties"
                    + "&depth=all"
                    + "&limit=" + limit
                    + "&start=" + start;

            JSONObject response = getJson(url);
            JSONArray results_ = response.getJSONArray("results");
            if (results_.isEmpty()) break;

            for (int i = 0; i < results_.length(); i++) {
                JSONObject c = results_.getJSONObject(i);

                JSONObject extensions = c.optJSONObject("extensions");
                if (extensions != null && extensions.has("inlineProperties")) continue;

                String id     = c.getString("id");
                String author = c.getJSONObject("history")
                        .getJSONObject("createdBy")
                        .getString("displayName");
                String created = c.getJSONObject("history")
                        .getString("createdDate");
                String updated = c.getJSONObject("version")
                        .getString("when");
                String bodyHtml = c.getJSONObject("body")
                        .getJSONObject("view")
                        .getString("value");
                String commentText = stripHtml(bodyHtml);

                String status = (extensions != null)
                        ? extensions.optString("resolution", "open")
                        : "open";

                String link = BASE_URL + "/pages/viewpage.action?pageId=" + pageId
                        + "&focusedCommentId=" + id + "#comment-" + id;

                CommentRecord rec = new CommentRecord();
                // Populate input metadata
                rec.module = inputRow.module;
                rec.legacyId = inputRow.legacyId;
                rec.newId = inputRow.newId;
                rec.functionName = inputRow.functionName;
                rec.functionType = inputRow.functionType;

                // Populate comment fields
                rec.type              = "Footer";
                rec.section           = "Page Level";
                rec.originalSelection = "";
                rec.id                = id;
                rec.author            = author;
                rec.createdDate       = created;
                rec.lastUpdated       = updated;
                rec.commentText       = commentText;
                rec.status            = status;
                rec.link              = link;
                results.add(rec);
            }

            JSONObject links = response.optJSONObject("_links");
            if (links == null || !links.has("next")) break;
            start += limit;
        }
        return results;
    }

    // -----------------------------------------------------------------------
    // Step 4 — write Excel
    // -----------------------------------------------------------------------
    /**
     * Write all collected CommentRecord objects into a single Excel workbook
     * and save it at the provided outputPath. The sheet includes input
     * metadata columns followed by comment details and a clickable link.
     *
     * The method also attempts to combine subsequent comments that lack an
     * "originalSelection" into the previous row's comment cell (appending the
     * text), preserving the original output structure.
     *
     * @param comments list of all comments to write into the workbook
     * @param outputPath destination .xlsx file path
     * @throws Exception on I/O or workbook write errors
     */
    private static void writeExcel(List<CommentRecord> comments, String outputPath)
            throws Exception {

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Comments");

            // Styling (keep existing styles)
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setBorderBottom(BorderStyle.THIN);

            CellStyle dataStyle = workbook.createCellStyle();
            dataStyle.setWrapText(true);
            dataStyle.setVerticalAlignment(VerticalAlignment.TOP);
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);

            CellStyle altStyle = workbook.createCellStyle();
            altStyle.cloneStyleFrom(dataStyle);
            altStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            altStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Font linkFont = workbook.createFont();
            linkFont.setColor(IndexedColors.BLUE.getIndex());
            linkFont.setUnderline(Font.U_SINGLE);

            CellStyle linkStyle = workbook.createCellStyle();
            linkStyle.cloneStyleFrom(dataStyle);
            linkStyle.setFont(linkFont);

            CellStyle linkAltStyle = workbook.createCellStyle();
            linkAltStyle.cloneStyleFrom(altStyle);
            linkAltStyle.setFont(linkFont);

            // Headers with input columns first
            String[] headers = {
                    "Module", "Legacy ID", "New ID", "Function Name", "Function Type",
                    "Type", "Section", "Highlighted Text",
                    "Comment ID", "Author", "Created Date", "Last Updated",
                    "Status", "Comment Text", "Link"
            };

            Row headerRow = sheet.createRow(0);
            headerRow.setHeightInPoints(20);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // Data rows
            int rowNum = 1;
            Row lastDataRow = null;

            for (CommentRecord rec : comments) {
                if ((rec.originalSelection == null || rec.originalSelection.trim().isEmpty()) && lastDataRow != null) {
                    // Append comment text to the previous row's comment cell (col 13)
                    Cell commentCell = lastDataRow.getCell(13);
                    if (commentCell != null) {
                        String existing = commentCell.getStringCellValue();
                        commentCell.setCellValue(
                                existing + (existing.isEmpty() ? "" : "\n") +
                                        (rec.commentText != null ? "[" + rec.createdDate + "]"
                                                + "[" + rec.author + "]) " + rec.commentText : "")
                        );
                    }
                    continue; // skip creating a new row
                }
                Row row = sheet.createRow(rowNum);
                row.setHeightInPoints(40);
                boolean alt = rowNum % 2 == 0;
                CellStyle cs    = alt ? altStyle    : dataStyle;
                CellStyle csLnk = alt ? linkAltStyle : linkStyle;

                // Input metadata columns
                createCell(row, 0, rec.module,            cs);
                createCell(row, 1, rec.legacyId,          cs);
                createCell(row, 2, rec.newId,             cs);
                createCell(row, 3, rec.functionName,      cs);
                createCell(row, 4, rec.functionType,      cs);

                // Comment columns
                createCell(row, 5, rec.type,              cs);
                createCell(row, 6, rec.section,           cs);
                createCell(row, 7, rec.originalSelection, cs);
                createCell(row, 8, rec.id,                cs);
                createCell(row, 9, rec.author,            cs);
                createCell(row, 10, rec.createdDate,      cs);
                createCell(row, 11, rec.lastUpdated,      cs);
                createCell(row, 12, rec.status,           cs);
                createCell(row, 13, "[" + rec.createdDate + "]"
                        + "[" + rec.author + "]) " + rec.commentText,cs);

                // Link cell
                Cell linkCell = row.createCell(14);
                linkCell.setCellValue("View Comment");
                linkCell.setCellStyle(csLnk);
                CreationHelper ch = workbook.getCreationHelper();
                Hyperlink hl = ch.createHyperlink(HyperlinkType.URL);
                hl.setAddress(rec.link);
                linkCell.setHyperlink(hl);

                lastDataRow = row;
                rowNum++;
            }

            // Column widths adjusted for more columns
            int[] widths = { 4000, 4000, 4000, 6000, 5000, 3000, 7000, 8000, 4000, 5000, 5500, 5500, 3500, 14000, 4000 };
            for (int i = 0; i < widths.length; i++) {
                sheet.setColumnWidth(i, widths[i]);
            }

            sheet.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(
                    0, rowNum - 1, 0, headers.length - 1));
            sheet.createFreezePane(0, 1);

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
            }
        }
        System.out.println("Excel written: " + outputPath);
    }

    /**
     * Create a cell at the specified column and set its string value and cell style.
     * This is a small convenience helper used throughout the workbook writer.
     *
     * @param row the row to create the cell on
     * @param col 0-based column index
     * @param value string value to set (null becomes empty string)
     * @param style CellStyle to apply
     */
    private static void createCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    /**
     * Strip HTML and return plain text. Uses Jsoup's HTML parser to remove tags
     * and decode HTML entities.
     *
     * @param html HTML string to strip
     * @return plain text content
     */
    private static String stripHtml(String html) {
        return Jsoup.parse(html).text();
    }

    /**
     * Perform an HTTP GET and parse the response body as JSON. The request
     * includes the configured Authorization header. On non-200 responses the
     * method throws a RuntimeException that includes the response body.
     *
     * @param urlStr fully-qualified URL to request
     * @return parsed JSONObject of the HTTP response body
     * @throws Exception on networking, HTTP or JSON parsing errors
     */
    private static JSONObject getJson(String urlStr) throws Exception {
        HttpClient client = HttpClient.newHttpClient();
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(urlStr))
                .header("Authorization", AUTH_HEADER)
                .header("Accept", "application/json")
                .GET()
                .build();
        HttpResponse<String> resp = client.send(request, HttpResponse.BodyHandlers.ofString());
        if (resp.statusCode() != 200) {
            throw new RuntimeException("HTTP " + resp.statusCode() + " for: " + urlStr
                    + "\nBody: " + resp.body());
        }
        return new JSONObject(resp.body());
    }

    // -----------------------------------------------------------------------
    // Data model
    // -----------------------------------------------------------------------
    /**
     * Represents one input row from the input Excel file. Holds module and
     * mapping metadata plus the Confluence pageId to process.
     */
    private static class InputRow {
        String module;
        String legacyId;
        String newId;
        String functionName;
        String functionType;
        String pageId;
    }

    /**
     * Represents a single comment with both input metadata and comment fields.
     * Instances are created while fetching inline/footer comments and are
     * eventually written to the output Excel file.
     */
    private static class CommentRecord {
        // Input metadata
        String module;
        String legacyId;
        String newId;
        String functionName;
        String functionType;

        // Comment fields
        String type;              // "Inline" | "Footer"
        String section;           // nearest heading or "Page Level"
        String originalSelection; // highlighted text (inline only)
        String id;
        String author;
        String createdDate;
        String lastUpdated;
        String commentText;
        String status;            // "open" | "resolved"
        String link;
    }

    /**
     * Summary row used for end-of-run reporting: maps input metadata to the
     * number of comments found on the pageId.
     */
    private static class SummaryRecord {
        String module;
        String legacyId;
        String newId;
        String functionName;
        String functionType;
        String pageId;
        int numberOfComments;
    }
}
