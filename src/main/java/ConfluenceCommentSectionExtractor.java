import org.apache.poi.ss.usermodel.*;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

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
    private static String PAGE_ID;
    
    private static String AUTH_HEADER;
    
    static {
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
    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            System.err.println("Usage: java ConfluenceCommentSectionExtractor <pageId>");
            System.exit(1);
        }
        PAGE_ID = args[0];
        String raw = EMAIL + ":" + API_TOKEN;
        AUTH_HEADER = "Basic " + Base64.getEncoder().encodeToString(raw.getBytes());
    
        System.out.println("=== Confluence Comment Section Extractor ===");
        System.out.println("Page ID : " + PAGE_ID);
        System.out.println("Base URL: " + BASE_URL);
        System.out.println();
        // 1. Fetch page body (storage format) to resolve section headings
        System.out.println("[1/3] Fetching page body (storage format)...");
        String storageBody = fetchPageBody();

        // 2. Build ordered list of (heading-text, list-of-inline-marker-refs) from body
        System.out.println("[2/3] Parsing headings and inline markers...");
        LinkedHashMap<String, List<String>> sectionMarkerMap = buildSectionMarkerMap(storageBody);
        // Reverse-lookup: markerRef -> sectionTitle
        Map<String, String> markerToSection = new HashMap<>();
        for (Map.Entry<String, List<String>> e : sectionMarkerMap.entrySet()) {
            for (String ref : e.getValue()) {
                markerToSection.put(ref, e.getKey());
            }
        }
        System.out.println("  Found " + sectionMarkerMap.size() + " section(s).");

        // 3. Fetch all comments (inline + footer)
        System.out.println("[3/3] Fetching comments...");
        List<CommentRecord> comments = new ArrayList<>();
        comments.addAll(fetchInlineComments(markerToSection));
        comments.addAll(fetchFooterComments());
        System.out.println("  Total comments fetched: " + comments.size());

        // 4. Write Excel
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String outputPath = "d:/outputs/confluence_comments_sections_" + timestamp + ".xlsx";
        writeExcel(comments, outputPath);

        System.out.println();
        System.out.println("Done! Output: " + outputPath);
    }

    // -----------------------------------------------------------------------
    // Step 1 — fetch page body in storage format
    // -----------------------------------------------------------------------
    private static String fetchPageBody() throws Exception {
        String url = BASE_URL + "/rest/api/content/" + PAGE_ID
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
    private static List<CommentRecord> fetchInlineComments(
            Map<String, String> markerToSection) throws Exception {

        List<CommentRecord> results = new ArrayList<>();
        int start = 0;
        int limit = 50;

        while (true) {
            String url = BASE_URL + "/rest/api/content/" + PAGE_ID
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

                // Only process inline comments (they have inlineProperties)
                JSONObject extensions = c.optJSONObject("extensions");
                if (extensions == null) continue;
                JSONObject inlineProps = extensions.optJSONObject("inlineProperties");
                if (inlineProps == null) continue; // footer comment — handled separately

                String markerRef        = inlineProps.optString("ref", "");
                String originalSelection = inlineProps.optString("originalSelection", "");

                String section = markerToSection.getOrDefault(markerRef, "(Section not resolved)");
                String id      = c.getString("id");
                String author  = c.getJSONObject("history")
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

                String status = extensions.optString("resolution", "open");

                String link = BASE_URL + "/pages/viewpage.action?pageId=" + PAGE_ID
                        + "&focusedCommentId=" + id + "#comment-" + id;

                CommentRecord rec = new CommentRecord();
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

            // Pagination
            JSONObject links = response.optJSONObject("_links");
            if (links == null || !links.has("next")) break;
            start += limit;
        }
        return results;
    }

    // -----------------------------------------------------------------------
    // Step 3b — fetch footer (page-level) comments
    // -----------------------------------------------------------------------
    private static List<CommentRecord> fetchFooterComments() throws Exception {
        List<CommentRecord> results = new ArrayList<>();
        int start = 0;
        int limit = 50;

        while (true) {
            String url = BASE_URL + "/rest/api/content/" + PAGE_ID
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
                // Skip inline comments (already handled)
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

                String link = BASE_URL + "/pages/viewpage.action?pageId=" + PAGE_ID
                        + "&focusedCommentId=" + id + "#comment-" + id;

                CommentRecord rec = new CommentRecord();
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
    private static void writeExcel(List<CommentRecord> comments, String outputPath)
            throws Exception {

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Comments");

            // ---------- styles ----------
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

            // ---------- header row ----------
            String[] headers = {
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

            // ---------- data rows ----------
            int rowNum = 1;
            for (CommentRecord rec : comments) {
                Row row = sheet.createRow(rowNum);
                row.setHeightInPoints(40);
                boolean alt = rowNum % 2 == 0;
                CellStyle cs    = alt ? altStyle    : dataStyle;
                CellStyle csLnk = alt ? linkAltStyle : linkStyle;

                createCell(row, 0, rec.type,              cs);
                createCell(row, 1, rec.section,           cs);
                createCell(row, 2, rec.originalSelection, cs);
                createCell(row, 3, rec.id,                cs);
                createCell(row, 4, rec.author,            cs);
                createCell(row, 5, rec.createdDate,       cs);
                createCell(row, 6, rec.lastUpdated,       cs);
                createCell(row, 7, rec.status,            cs);
                createCell(row, 8, rec.commentText,       cs);

                // Hyperlink cell
                Cell linkCell = row.createCell(9);
                linkCell.setCellValue("View Comment");
                linkCell.setCellStyle(csLnk);
                CreationHelper ch = workbook.getCreationHelper();
                Hyperlink hl = ch.createHyperlink(HyperlinkType.URL);
                hl.setAddress(rec.link);
                linkCell.setHyperlink(hl);

                rowNum++;
            }

            // ---------- column widths ----------
            int[] widths = { 3000, 7000, 8000, 4000, 5000, 5500, 5500, 3500, 14000, 4000 };
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

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------
    private static void createCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private static String stripHtml(String html) {
        return Jsoup.parse(html).text();
    }

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
    static class CommentRecord {
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
}
