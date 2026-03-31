import org.apache.poi.ss.usermodel.*;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.*;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Base64;
import java.util.Base64;
import java.util.regex.*;

/**
 * Reads a Jira export Excel file (jira-search-result.xlsx) and produces
 * an output Excel report with the following columns:
 *
 *   Module, Legacy ID, New ID, Function Name, Function Type,
 *   Page ID, Reviewer Name, Confluence URL, Full URL
 *
 * Field derivation rules:
 *   Module        = chars [1..3] of New Function ID (e.g. WCMN00190 → CMN)
 *   Function Type = first char of New Function ID:
 *                     W → Screen | B → Batch | R → Report | I → Interface
 *   Page ID       = extracted from Confluence DR Doc URL (handles full URLs,
 *                   edit-v2 URLs, viewpage.action, resumedraft.action, /wiki/x/ tiny URLs)
 *
 * pom.xml dependencies required:
 *   - org.apache.poi : poi-ooxml    (for Excel read/write)
 *
 * Usage (standalone):
 *   java -cp ... JiraDRReportGenerator path/to/jira-search-result.xlsx
 *
 * Usage (programmatic):
 *   JiraDRReportGenerator gen = JiraDRReportGenerator.fromConfig();
 *   gen.generate(Paths.get("jira-search-result.xlsx"), Paths.get("output_dir"));
 */
public class JiraDRReportGenerator {

    // ─── Input column names (from 'Your Jira Issues' sheet) ───────────────────
    private static final String COL_KEY          = "Key";
    private static final String COL_SUMMARY      = "Summary";
    private static final String COL_LEGACY_ID    = "Legacy Function ID";
    private static final String COL_NEW_ID       = "New Function ID";
    private static final String COL_CONFLUENCE   = "Confluence DR Doc";
    private static final String COL_DR_REVIEWER  = "DR_Reviewer";
    private static final String DATA_SHEET       = "Your Jira Issues";

    // ─── Output column headers ─────────────────────────────────────────────────
    private static final String[] OUTPUT_HEADERS = {
        "Module", "Legacy ID", "New ID", "Function Name",
        "Function Type", "Page ID", "Reviewer Name",
        "Confluence URL", "Full URL", "Jira Ticket ID"
    };

    // ─── URL page ID extraction patterns (tried in order) ─────────────────────
    private static final List<Pattern> PAGE_ID_PATTERNS = List.of(
        Pattern.compile("/pages/(\\d+)(?:/|$|\\?)"),       // /pages/{id}/Title
        Pattern.compile("/pages/edit-v2/(\\d+)"),           // /pages/edit-v2/{id}
        Pattern.compile("[?&]pageId=(\\d+)"),               // ?pageId={id}
        Pattern.compile("[?&]draftId=(\\d+)")               // ?draftId={id}
    );

    private static final int    MAX_REDIRECT_HOPS = 10;
    private static final String CONFIG_FILE       = "config.properties";
    private static final String PROP_TENANT_URL   = "tenant.url";
    private static final String PROP_EMAIL        = "email";
    private static final String PROP_API_TOKEN    = "api.token";

    // ─── Output styling constants ──────────────────────────────────────────────
    private static final String HEADER_COLOR = "1F4E79";
    private static final String ALT_ROW_COLOR = "EBF2FA";

    private final String     tenantUrl;
    private final String     email;
    private final String     apiToken;
    private final HttpClient httpClient;

    // ─── Constructor ───────────────────────────────────────────────────────────

    public JiraDRReportGenerator(String tenantUrl, String email, String apiToken) {
        if (tenantUrl == null || tenantUrl.isBlank())
            throw new IllegalArgumentException("tenantUrl must not be blank");
        if (email == null || email.isBlank())
            throw new IllegalArgumentException("email must not be blank");
        if (apiToken == null || apiToken.isBlank())
            throw new IllegalArgumentException("apiToken must not be blank");
        this.tenantUrl  = tenantUrl.replaceAll("/+$", "");
        this.email      = email;
        this.apiToken   = apiToken;
        // Redirects disabled — we follow the chain manually in resolveConfluenceTinyUrl
        this.httpClient = HttpClient.newBuilder()
                .followRedirects(HttpClient.Redirect.NEVER)
                .connectTimeout(Duration.ofSeconds(10))
                .build();
    }

    /** Factory: reads jira.tenant.url from config.properties on classpath. */
    public static JiraDRReportGenerator fromConfig() {
        Properties props = new Properties();
        try (java.io.InputStream is = new java.io.FileInputStream(CONFIG_FILE)) {
            if (is == null)
                throw new IllegalStateException(CONFIG_FILE + " not found on classpath.");
            props.load(is);
        } catch (IOException e) {
            throw new IllegalStateException("Failed to read " + CONFIG_FILE, e);
        }
        String url      = requireProp(props, PROP_TENANT_URL);
        String email    = requireProp(props, PROP_EMAIL);
        String apiToken = requireProp(props, PROP_API_TOKEN);
        return new JiraDRReportGenerator(url, email, apiToken);
    }

    private static String requireProp(Properties props, String key) {
        String v = props.getProperty(key);
        if (v == null || v.isBlank())
            throw new IllegalStateException(
                "Missing property '" + key + "' in " + CONFIG_FILE);
        return v.trim();
    }

    // ─── Public API ────────────────────────────────────────────────────────────

    /**
     * Reads inputFile, resolves Confluence page IDs, and writes output Excel.
     *
     * @param inputFile  path to jira-search-result.xlsx
     * @param outputDir  directory where output_yyyyMMddHHmmss.xlsx will be written
     * @return path to the generated output file
     */
    public Path generate(Path inputFile, Path outputDir) throws IOException {
        System.out.printf("Reading: %s%n", inputFile);
        List<Map<String, String>> dataRows = readInputExcel(inputFile);
        System.out.printf("Processing %d rows...%n", dataRows.size());

        List<OutputRow> outputRows = new ArrayList<>();
        int i = 0;
        for (Map<String, String> row : dataRows) {
            i++;
            OutputRow out = processRow(row);
            outputRows.add(out);
            System.out.printf("  [%4d/%d] %-12s %-12s pageId=%-14s%n",
                i, dataRows.size(), row.getOrDefault(COL_KEY, ""),
                out.newId, out.pageId.isEmpty() ? "?" : out.pageId);
        }

        String timestamp  = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
        Path   outputPath = outputDir.resolve("output_" + timestamp + ".xlsx");
        Files.createDirectories(outputDir);
        writeOutputExcel(outputRows, outputPath);

        long resolved = outputRows.stream().filter(r -> !r.pageId.isEmpty()).count();
        System.out.printf("%nDone → %s%n", outputPath);
        System.out.printf("Page IDs resolved: %d / %d%n", resolved, outputRows.size());
        return outputPath;
    }

    // ─── Private: Input reading ────────────────────────────────────────────────

    private List<Map<String, String>> readInputExcel(Path path) throws IOException {
        List<Map<String, String>> rows = new ArrayList<>();
        // Disable POI zip bomb detection — Jira export files have a high
        // compression ratio that triggers the default 0.01 threshold.
        // Setting to 0 disables the check entirely (safe for trusted input).
        ZipSecureFile.setMinInflateRatio(0);
        try (Workbook wb = WorkbookFactory.create(path.toFile())) {
            Sheet sheet = wb.getSheet(DATA_SHEET);
            if (sheet == null)
                throw new IllegalArgumentException(
                    "Sheet '" + DATA_SHEET + "' not found in: " + path);

            // Build column index map from header row
            Row header = sheet.getRow(0);
            Map<String, Integer> colIndex = new LinkedHashMap<>();
            for (Cell cell : header) {
                String name = getCellString(cell);
                if (!name.isBlank()) colIndex.put(name, cell.getColumnIndex());
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Map<String, String> rowData = new LinkedHashMap<>();
                for (Map.Entry<String, Integer> e : colIndex.entrySet()) {
                    Cell cell = row.getCell(e.getValue());
                    rowData.put(e.getKey(), cell != null ? getCellString(cell) : "");
                }
                rows.add(rowData);
            }
        }
        return rows;
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING  -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                yield (v == Math.floor(v)) ? String.valueOf((long) v) : String.valueOf(v);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try { yield String.valueOf(cell.getStringCellValue()).trim(); }
                catch (Exception ex) { yield String.valueOf(cell.getNumericCellValue()); }
            }
            default -> "";
        };
    }

    // ─── Private: Row processing ───────────────────────────────────────────────

    private OutputRow processRow(Map<String, String> row) {
        String newId      = row.getOrDefault(COL_NEW_ID, "").trim();
        String legacyId   = row.getOrDefault(COL_LEGACY_ID, "").trim();
        String summary    = row.getOrDefault(COL_SUMMARY, "").trim();
        String reviewer   = row.getOrDefault(COL_DR_REVIEWER, "").trim();
        String confUrl    = row.getOrDefault(COL_CONFLUENCE, "").trim();
        String jiraTicketId = row.getOrDefault(COL_KEY, "").trim();

        String module       = deriveModule(newId);
        String functionType = deriveFunctionType(newId);

        // Step 1: try to extract page ID directly from the URL (no HTTP needed)
        String pageId  = extractPageId(confUrl);
        String fullUrl = confUrl;

        // Step 2: if page ID still missing, resolve via HTTP redirect chain.
        //         This covers ALL unresolvable URL types:
        //           - Tiny URLs          /wiki/x/{code}
        //           - Resume draft URLs  /wiki/pages/resumedraft.action?draftId={id}
        //           - Viewpage URLs      /wiki/pages/viewpage.action?pageId={id}
        //           - Any other URL that does not contain a numeric page ID
        if (pageId == null && !confUrl.isBlank()) {
            ResolvedUrl resolved = resolveUrl(confUrl);
            // Only update fullUrl when we actually resolved to a different URL
            if (resolved.url != null && !resolved.url.equals(confUrl)) {
                fullUrl = resolved.url;
            }
            pageId = resolved.pageId;
        }

        return new OutputRow(
            module, legacyId, newId, summary, functionType,
            pageId  != null ? pageId  : "",
            reviewer, confUrl,
            fullUrl != null ? fullUrl : "",
            jiraTicketId
        );
    }

    private String deriveModule(String newId) {
        if (newId == null || newId.length() < 4) return "";
        return newId.substring(1, 4);  // chars at index 1,2,3
    }

    private String deriveFunctionType(String newId) {
        if (newId == null || newId.isEmpty()) return "";
        return switch (newId.charAt(0)) {
            case 'W', 'w' -> "Screen";
            case 'B', 'b' -> "Batch";
            case 'R', 'r' -> "Report";
            case 'I', 'i' -> "Interface";
            default        -> "";
        };
    }

    private String extractPageId(String url) {
        if (url == null || url.isBlank()) return null;
        // Strip query string for pattern matching on path only, but also check full
        for (Pattern p : PAGE_ID_PATTERNS) {
            Matcher m = p.matcher(url);
            if (m.find()) return m.group(1);
        }
        return null;
    }

    // ─── Private: URL resolution ──────────────────────────────────────────────

    private record ResolvedUrl(String url, String pageId) {}

    /**
     * Resolves any Confluence URL to its final full URL and extracts the page ID.
     *
     * Uses the same authenticated redirect-following logic as JiraDRInfoFetcher
     * so that short URLs (/wiki/x/), resumedraft, viewpage.action, etc.
     * are all handled correctly with Basic Auth credentials.
     *
     * Confluence short URLs go through a two-hop redirect chain:
     *   Hop 1: /wiki/x/{code}
     *       ->  /wiki/pages/tinyurl.action?urlIdentifier={code}
     *   Hop 2: /wiki/pages/tinyurl.action?urlIdentifier={code}
     *       ->  /wiki/spaces/{space}/pages/{id}/Title
     */
    private ResolvedUrl resolveUrl(String url) {
        String currentUrl = url;

        for (int hop = 1; hop <= MAX_REDIRECT_HOPS; hop++) {
            System.out.printf("  [redirect hop %d] GET %s%n", hop, currentUrl);
            HttpResponse<String> response = doGet(currentUrl);
            int status = response.statusCode();

            if (status == 301 || status == 302 || status == 303
                    || status == 307 || status == 308) {

                String location = response.headers().firstValue("Location").orElse(null);
                if (location == null) {
                    System.err.printf("  WARN: Redirect hop %d had no Location header for: %s%n",
                        hop, currentUrl);
                    break;
                }

                // Make relative redirects absolute
                if (location.startsWith("/")) {
                    location = tenantUrl + location;
                }

                System.out.printf("  [redirect hop %d] -> %s%n", hop, location);
                currentUrl = location;

                // Check full URL including query string after each hop
                // (covers /pages/{id}/, ?pageId=, ?draftId=, edit-v2/{id})
                String pid = extractPageId(currentUrl);
                if (pid != null) {
                    System.out.printf("  [redirect hop %d] Found page ID %s -- stopping.%n",
                        hop, pid);
                    return new ResolvedUrl(currentUrl, pid);
                }

            } else if (status == 200) {
                // Final destination -- try to extract page ID from wherever we ended up
                String pid = extractPageId(currentUrl);
                if (pid == null) {
                    System.err.printf(
                        "WARN: Redirect chain ended at HTTP 200 without a page ID: %s%n",
                        currentUrl);
                }
                return new ResolvedUrl(currentUrl, pid);

            } else if (status == 401 || status == 403) {
                System.err.printf("  WARN: Authentication failed (HTTP %d) for: %s%n",
                    status, currentUrl);
                System.err.println("        Check jira.email and jira.api.token in config.properties");
                break;

            } else {
                System.err.printf("  WARN: Unexpected HTTP %d at hop %d for: %s%n",
                    status, hop, currentUrl);
                break;
            }
        }

        return new ResolvedUrl(url, null);
    }

    /**
     * Executes an authenticated GET request using Atlassian Basic Auth (email:apiToken).
     * Same implementation as JiraDRInfoFetcher.doGet().
     */
    private HttpResponse<String> doGet(String url) {
        String credentials = Base64.getEncoder()
                .encodeToString((email + ":" + apiToken).getBytes());

        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Basic " + credentials)
                .header("Accept", "text/html,application/json")
                .timeout(Duration.ofSeconds(15))
                .GET()
                .build();

        try {
            return httpClient.send(request, HttpResponse.BodyHandlers.ofString());
        } catch (IOException e) {
            throw new RuntimeException("Network error calling: " + url, e);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new RuntimeException("Request interrupted for: " + url, e);
        }
    }

    // ─── Private: Output Excel writing ────────────────────────────────────────

    private static final int[] COL_WIDTHS_CHARS = {10, 14, 14, 55, 14, 16, 22, 50, 70, 15};

    private void writeOutputExcel(List<OutputRow> rows, Path outputPath) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("DR Report");

            // Styles
            XSSFCellStyle headerStyle = createHeaderStyle(wb);
            XSSFCellStyle centerStyle = createDataStyle(wb, false, false);
            XSSFCellStyle leftStyle   = createDataStyle(wb, false, true);
            XSSFCellStyle linkStyle   = createDataStyle(wb, true,  true);
            XSSFCellStyle centerAlt   = createDataStyle(wb, false, false);
            XSSFCellStyle leftAlt     = createDataStyle(wb, false, true);
            XSSFCellStyle linkAlt     = createDataStyle(wb, true,  true);
            applyAltFill(wb, centerAlt, linkAlt, leftAlt);

            // Header row
            Row header = ws.createRow(0);
            header.setHeightInPoints(30);
            for (int ci = 0; ci < OUTPUT_HEADERS.length; ci++) {
                Cell cell = header.createCell(ci);
                cell.setCellValue(OUTPUT_HEADERS[ci]);
                cell.setCellStyle(headerStyle);
                ws.setColumnWidth(ci, COL_WIDTHS_CHARS[ci] * 256);
            }
            ws.createFreezePane(0, 1);
            ws.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(
                0, rows.size(), 0, OUTPUT_HEADERS.length - 1));

            // Data rows
            for (int ri = 0; ri < rows.size(); ri++) {
                OutputRow r    = rows.get(ri);
                Row       row  = ws.createRow(ri + 1);
                boolean   isAlt = (ri + 1) % 2 == 0;

                String[] values = {
                    r.module, r.legacyId, r.newId, r.functionName,
                    r.functionType, r.pageId, r.reviewerName,
                    r.confluenceUrl, r.fullUrl, r.jiraTicketId
                };

                for (int ci = 0; ci < values.length; ci++) {
                    Cell cell = row.createCell(ci);
                    cell.setCellValue(values[ci]);
                    boolean isUrl = ci >= 7 && values[ci].startsWith("http");
                    boolean isName = ci == 3;
                    if (isUrl) {
                        if (isAlt) { cell.setCellStyle(linkAlt); }
                        else       { cell.setCellStyle(linkStyle); }
                        if (!values[ci].isBlank())
                            cell.setHyperlink(createHyperlink(wb, values[ci]));
                    } else if (isName) {
                        cell.setCellStyle(isAlt ? leftAlt : leftStyle);
                    } else {
                        cell.setCellStyle(isAlt ? centerAlt : centerStyle);
                    }
                }
            }

            try (OutputStream os = Files.newOutputStream(outputPath)) {
                wb.write(os);
            }
        }
    }

    private XSSFCellStyle createHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(hexToBytes(HEADER_COLOR), null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        setBorder(style, BorderStyle.THIN);
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(new byte[]{(byte)0xFF, (byte)0xFF, (byte)0xFF}, null));
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        return style;
    }

    private XSSFCellStyle createDataStyle(XSSFWorkbook wb, boolean isLink, boolean leftAlign) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setAlignment(leftAlign ? HorizontalAlignment.LEFT : HorizontalAlignment.CENTER);
        style.setWrapText(leftAlign && !isLink);
        setBorder(style, BorderStyle.THIN);
        XSSFFont font = wb.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        if (isLink) {
            font.setColor(new XSSFColor(new byte[]{0x05, 0x63, (byte)0xC1}, null));
            font.setUnderline(Font.U_SINGLE);
        }
        style.setFont(font);
        return style;
    }

    private void applyAltFill(XSSFWorkbook wb, XSSFCellStyle... styles) {
        for (XSSFCellStyle s : styles) {
            s.setFillForegroundColor(new XSSFColor(hexToBytes(ALT_ROW_COLOR), null));
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
    }

    private void setBorder(XSSFCellStyle style, BorderStyle bs) {
        style.setBorderTop(bs); style.setBorderBottom(bs);
        style.setBorderLeft(bs); style.setBorderRight(bs);
    }

    private Hyperlink createHyperlink(XSSFWorkbook wb, String url) {
        CreationHelper ch   = wb.getCreationHelper();
        Hyperlink      link = ch.createHyperlink(HyperlinkType.URL);
        link.setAddress(url);
        return link;
    }

    private static byte[] hexToBytes(String hex) {
        int len = hex.length();
        byte[] data = new byte[len / 2];
        for (int i = 0; i < len; i += 2)
            data[i / 2] = (byte) ((Character.digit(hex.charAt(i), 16) << 4)
                                 + Character.digit(hex.charAt(i + 1), 16));
        return data;
    }

    // ─── Output Row model ──────────────────────────────────────────────────────

    private record OutputRow(
        String module, String legacyId, String newId,
        String functionName, String functionType,
        String pageId, String reviewerName,
        String confluenceUrl, String fullUrl, String jiraTicketId
    ) {}

    // ─── Main ──────────────────────────────────────────────────────────────────

    public static void main(String[] args) throws IOException {
        if (args.length < 1) {
            System.err.println("Usage: JiraDRReportGenerator <input.xlsx> [output-dir]");
            System.exit(1);
        }
        Path inputFile = Paths.get(args[0]);
        Path outputDir = args.length > 1 ? Paths.get(args[1]) : inputFile.getParent();

        JiraDRReportGenerator generator;
        try {
            generator = JiraDRReportGenerator.fromConfig();
        } catch (Exception e) {
            System.err.println("Config error: " + e.getMessage());
            System.exit(1);
            return;
        }
        generator.generate(inputFile, outputDir);
    }
}
