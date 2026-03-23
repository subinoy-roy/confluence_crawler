import java.util.Properties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;

public class ConfluenceCommentsExtractor {
    // -----------------------------------------------------------------------
    // Configuration — update these before running
    // -----------------------------------------------------------------------
    private static final String BASE_URL;
    private static final String EMAIL;
    private static final String API_TOKEN;

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

    public static void main(String[] args) {
        String PAGE_ID = args[0];
        String raw = EMAIL + ":" + API_TOKEN;

        try {
            System.out.println("Starting Confluence comments extraction...");

            List<Comment> comments = extractComments(BASE_URL, PAGE_ID, EMAIL, API_TOKEN);

            if (comments != null && !comments.isEmpty()) {
                String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
                String outputFile = "confluence_comments_" + timestamp + ".xlsx";

                createExcelReport(comments, outputFile);
                System.out.println("Successfully created: " + outputFile);
            } else {
                System.out.println("No comments found or error occurred.");
            }

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static List<Comment> extractComments(String baseUrl, String pageId, String email, String apiToken) throws IOException {
        List<Comment> comments = new ArrayList<>();
        int start = 0;
        int limit = 100;

        String auth = email + ":" + apiToken;
        String encodedAuth = Base64.getEncoder().encodeToString(auth.getBytes(StandardCharsets.UTF_8));

        System.out.println("Fetching comments from page ID: " + pageId);

        while (true) {
            String urlString = String.format("%s/rest/api/content/%s/child/comment?expand=body.view,version,history.lastUpdated&start=%d&limit=%d",
                    baseUrl, pageId, start, limit);

            URL url = new URL(urlString);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Authorization", "Basic " + encodedAuth);
            conn.setRequestProperty("Accept", "application/json");

            int responseCode = conn.getResponseCode();

            if (responseCode != 200) {
                System.err.println("HTTP Error: " + responseCode);
                BufferedReader errorReader = new BufferedReader(new InputStreamReader(conn.getErrorStream()));
                String line;
                while ((line = errorReader.readLine()) != null) {
                    System.err.println(line);
                }
                errorReader.close();
                return null;
            }

            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            StringBuilder response = new StringBuilder();
            String line;
            while ((line = in.readLine()) != null) {
                response.append(line);
            }
            in.close();

            JSONObject jsonResponse = new JSONObject(response.toString());
            JSONArray results = jsonResponse.getJSONArray("results");

            if (results.length() == 0) {
                break;
            }

            for (int i = 0; i < results.length(); i++) {
                JSONObject commentJson = results.getJSONObject(i);
                Comment comment = new Comment();

                comment.id = commentJson.optString("id", "");

                if (commentJson.has("history") && commentJson.getJSONObject("history").has("createdBy")) {
                    comment.author = commentJson.getJSONObject("history")
                            .getJSONObject("createdBy")
                            .optString("displayName", "Unknown");
                }

                if (commentJson.has("history")) {
                    comment.createdDate = commentJson.getJSONObject("history").optString("createdDate", "");
                }

                if (commentJson.has("version")) {
                    comment.lastUpdated = commentJson.getJSONObject("version").optString("when", "");
                }

                if (commentJson.has("body") && commentJson.getJSONObject("body").has("view")) {
                    String htmlContent = commentJson.getJSONObject("body")
                            .getJSONObject("view")
                            .optString("value", "");
                    comment.commentText = cleanHtml(htmlContent);
                }

                // Construct comment link
                comment.commentLink = String.format("%s/pages/viewpage.action?pageId=%s&focusedCommentId=%s#comment-%s",
                        baseUrl, pageId, comment.id, comment.id);

                comments.add(comment);
            }

            if (results.length() < limit) {
                break;
            }

            start += limit;
        }

        System.out.println("Total comments extracted: " + comments.size());
        return comments;
    }

    private static String cleanHtml(String html) {
        if (html == null || html.isEmpty()) {
            return "";
        }

        String cleaned = html.replaceAll("<[^>]*>", "");
        cleaned = cleaned.replaceAll("&nbsp;", " ");
        cleaned = cleaned.replaceAll("&lt;", "<");
        cleaned = cleaned.replaceAll("&gt;", ">");
        cleaned = cleaned.replaceAll("&amp;", "&");
        cleaned = cleaned.replaceAll("&quot;", "\"");

        return cleaned.trim();
    }

    private static void createExcelReport(List<Comment> comments, String outputFile) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Comments");

        // Create header style
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font headerFont = workbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Create data style
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setWrapText(true);
        dataStyle.setVerticalAlignment(VerticalAlignment.TOP);

        // Create header row
        Row headerRow = sheet.createRow(0);
        String[] headers = {"Comment ID", "Author", "Created Date", "Last Updated", "Comment Text", "Comment Link"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Add data rows
        int rowNum = 1;
        for (Comment comment : comments) {
            Row row = sheet.createRow(rowNum++);

            Cell cell0 = row.createCell(0);
            cell0.setCellValue(comment.id);
            cell0.setCellStyle(dataStyle);

            Cell cell1 = row.createCell(1);
            cell1.setCellValue(comment.author);
            cell1.setCellStyle(dataStyle);

            Cell cell2 = row.createCell(2);
            cell2.setCellValue(comment.createdDate);
            cell2.setCellStyle(dataStyle);

            Cell cell3 = row.createCell(3);
            cell3.setCellValue(comment.lastUpdated);
            cell3.setCellStyle(dataStyle);

            Cell cell4 = row.createCell(4);
            cell4.setCellValue(comment.commentText);
            cell4.setCellStyle(dataStyle);

            // Add hyperlink for Comment Link
            Cell cell5 = row.createCell(5);
            cell5.setCellValue("View Comment");

            // Create hyperlink
            CreationHelper createHelper = workbook.getCreationHelper();
            Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(comment.commentLink);
            cell5.setHyperlink(link);

            // Style for hyperlink
            CellStyle linkStyle = workbook.createCellStyle();
            linkStyle.setVerticalAlignment(VerticalAlignment.TOP);
            Font linkFont = workbook.createFont();
            linkFont.setUnderline(Font.U_SINGLE);
            linkFont.setColor(IndexedColors.BLUE.getIndex());
            linkStyle.setFont(linkFont);
            cell5.setCellStyle(linkStyle);
        }

        // Set column widths
        sheet.setColumnWidth(0, 4000);  // Comment ID
        sheet.setColumnWidth(1, 5000);  // Author
        sheet.setColumnWidth(2, 5000);  // Created Date
        sheet.setColumnWidth(3, 5000);  // Last Updated
        sheet.setColumnWidth(4, 15000); // Comment Text
        sheet.setColumnWidth(5, 4000);  // Comment Link

        // Enable auto-filter
        sheet.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(0, rowNum - 1, 0, headers.length - 1));

        // Write to file
        try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
            workbook.write(outputStream);
        }

        workbook.close();
    }

    static class Comment {
        String id;
        String author;
        String createdDate;
        String lastUpdated;
        String commentText;
        String commentLink;
    }
}