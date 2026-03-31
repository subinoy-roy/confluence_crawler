import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Fetches DR Reviewer, DR Creator, Confluence DR Doc URL,
 * Confluence Page ID, and Confluence Page Title from a Jira ticket.
 *
 * ─── Why two different base URLs? ────────────────────────────────────────────
 * Atlassian Cloud exposes its APIs via two different hosts:
 *
 *   Jira REST API  → https://api.atlassian.com/ex/jira/{cloudId}/rest/api/3/
 *   Confluence API → https://{tenant}.atlassian.net/wiki/rest/api/
 *
 * Using the tenant URL for Jira returns 404. Always use api.atlassian.com.
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * Configuration is loaded from config.properties on the classpath:
 *
 *   jira.cloud.id    = 0d623cb3-b0f5-4952-8810-f649723c7b67
 *   jira.tenant.url  = https://tdem.atlassian.net
 *   jira.email       = your.email@company.com
 *   jira.api.token   = your_api_token
 *
 * Custom field mappings (specific to the DM project):
 *   customfield_12491 → DR_Reviewer  (multi-user)
 *   customfield_12493 → DR_Creator   (multi-user)
 *   customfield_12062 → Confluence DR Doc URL (tiny URL string)
 */
public class JiraDRInfoFetcher {

    // ─── Config property keys ──────────────────────────────────────────────────
    private static final String CONFIG_FILE      = "config.properties";
    private static final String PROP_CLOUD_ID    = "cloud.id";
    private static final String PROP_TENANT_URL  = "tenant.url";
    private static final String PROP_EMAIL       = "email";
    private static final String PROP_API_TOKEN   = "api.token";

    // ─── Custom Field IDs ──────────────────────────────────────────────────────
    private static final String FIELD_DR_REVIEWER    = "customfield_12491";
    private static final String FIELD_DR_CREATOR     = "customfield_12493";
    private static final String FIELD_CONFLUENCE_URL = "customfield_12062";

    // Matches /pages/{numeric-id}/ in a full Confluence URL
    private static final Pattern PAGE_ID_PATTERN =
            Pattern.compile("/pages/(\\d+)(?:/|$)");

    // ─── Instance fields ───────────────────────────────────────────────────────

    /** https://api.atlassian.com/ex/jira/{cloudId} — correct host for Jira Cloud */
    private final String jiraApiBase;

    /** https://{tenant}.atlassian.net — used for Confluence API + tiny URL resolution */
    private final String tenantUrl;

    private final String email;
    private final String apiToken;

    private final HttpClient  httpClient;
    private final ObjectMapper mapper;

    // ─── Result Model ──────────────────────────────────────────────────────────

    public static class DRInfo {
        public final String       ticketId;
        public final List<String> drReviewers;
        public final List<String> drCreators;
        public final String       confluenceTinyUrl;
        public final String       confluenceFullUrl;
        public final String       confluencePageId;
        public final String       confluencePageTitle;

        public DRInfo(String ticketId,
                      List<String> drReviewers,
                      List<String> drCreators,
                      String confluenceTinyUrl,
                      String confluenceFullUrl,
                      String confluencePageId,
                      String confluencePageTitle) {
            this.ticketId            = ticketId;
            this.drReviewers         = drReviewers;
            this.drCreators          = drCreators;
            this.confluenceTinyUrl   = confluenceTinyUrl;
            this.confluenceFullUrl   = confluenceFullUrl;
            this.confluencePageId    = confluencePageId;
            this.confluencePageTitle = confluencePageTitle;
        }

        @Override
        public String toString() {
            return String.format(
                "┌─ %s ───────────────────────────────────────%n" +
                "│  DR Reviewer(s)  : %s%n" +
                "│  DR Creator(s)   : %s%n" +
                "│  Confluence URL  : %s%n" +
                "│  Full URL        : %s%n" +
                "│  Page ID         : %s%n" +
                "│  Page Title      : %s%n" +
                "└──────────────────────────────────────────────",
                ticketId,
                drReviewers.isEmpty() ? "(none)" : String.join(", ", drReviewers),
                drCreators.isEmpty()  ? "(none)" : String.join(", ", drCreators),
                confluenceTinyUrl   != null ? confluenceTinyUrl   : "(not set)",
                confluenceFullUrl   != null ? confluenceFullUrl   : "(not resolved)",
                confluencePageId    != null ? confluencePageId    : "(not found)",
                confluencePageTitle != null ? confluencePageTitle : "(not found)"
            );
        }
    }

    // ─── Custom Exceptions ─────────────────────────────────────────────────────

    public static class JiraDRException extends RuntimeException {
        public JiraDRException(String message)                  { super(message); }
        public JiraDRException(String message, Throwable cause) { super(message, cause); }
    }

    public static class ConfigurationException extends JiraDRException {
        public ConfigurationException(String message) { super(message); }
    }

    public static class JiraIssueNotFoundException extends JiraDRException {
        public JiraIssueNotFoundException(String ticketId) {
            super("Jira issue not found: " + ticketId);
        }
    }

    public static class JiraAuthException extends JiraDRException {
        public JiraAuthException(String context) {
            super("Authentication failed (" + context + "). " +
                  "Check jira.email and jira.api.token in " + CONFIG_FILE);
        }
    }

    public static class ConfluencePageNotFoundException extends JiraDRException {
        public ConfluencePageNotFoundException(String pageId) {
            super("Confluence page not found for ID: " + pageId);
        }
    }

    // ─── Constructors ──────────────────────────────────────────────────────────

    /**
     * Primary constructor — accepts config values directly.
     *
     * @param cloudId   Atlassian Cloud ID UUID
     * @param tenantUrl Full tenant URL, e.g. https://tdem.atlassian.net
     * @param email     Atlassian account email
     * @param apiToken  Atlassian API token
     */
    public JiraDRInfoFetcher(String cloudId, String tenantUrl,
                             String email, String apiToken) {
        requireNonBlank(cloudId,   PROP_CLOUD_ID);
        requireNonBlank(tenantUrl, PROP_TENANT_URL);
        requireNonBlank(email,     PROP_EMAIL);
        requireNonBlank(apiToken,  PROP_API_TOKEN);

        // Jira Cloud API must go through api.atlassian.com — NOT the tenant URL
        this.jiraApiBase = "https://api.atlassian.com/ex/jira/" + cloudId;
        this.tenantUrl   = tenantUrl.replaceAll("/+$", "");
        this.email       = email;
        this.apiToken    = apiToken;
        this.mapper      = new ObjectMapper();

        // Redirect-disabled client — needed to capture Location header
        // when resolving Confluence /wiki/x/ tiny URLs
        this.httpClient  = HttpClient.newBuilder()
                .followRedirects(HttpClient.Redirect.NEVER)
                .connectTimeout(Duration.ofSeconds(10))
                .build();
    }

    /**
     * Factory method — loads configuration from config.properties on the classpath.
     *
     * Expected file location: src/main/resources/config.properties
     *
     * @throws ConfigurationException if the file is missing or any required key is absent
     */
    public static JiraDRInfoFetcher fromConfig() {
        Properties props = loadProperties();
        return new JiraDRInfoFetcher(
            requireProp(props, PROP_CLOUD_ID),
            requireProp(props, PROP_TENANT_URL),
            requireProp(props, PROP_EMAIL),
            requireProp(props, PROP_API_TOKEN)
        );
    }

    // ─── Public API ────────────────────────────────────────────────────────────

    /**
     * Fetches DR information for a single Jira ticket.
     *
     * @param ticketId  e.g. "DM-2199"
     * @return populated DRInfo record
     * @throws JiraDRException on any unrecoverable error
     */
    public DRInfo fetch(String ticketId) {
        requireNonBlank(ticketId, "ticketId");
        String key = ticketId.toUpperCase().trim();

        System.out.printf("[%s] Fetching Jira issue...%n", key);
        JsonNode fields = fetchJiraIssueFields(key);

        List<String> reviewers = extractUserList(fields, FIELD_DR_REVIEWER);
        List<String> creators  = extractUserList(fields, FIELD_DR_CREATOR);
        String tinyUrl         = extractStringField(fields, FIELD_CONFLUENCE_URL);

        String fullUrl   = null;
        String pageId    = null;
        String pageTitle = null;

        if (tinyUrl != null) {
            System.out.printf("[%s] Resolving Confluence tiny URL: %s%n", key, tinyUrl);
            fullUrl = resolveConfluenceTinyUrl(tinyUrl);

            if (fullUrl != null) {
                pageId = extractPageIdFromUrl(fullUrl);
            }

            if (pageId != null) {
                System.out.printf("[%s] Fetching Confluence page title for ID: %s%n", key, pageId);
                pageTitle = fetchConfluencePageTitle(pageId, key);
            } else {
                System.err.printf(
                    "[%s] WARN: Could not extract page ID from URL: %s%n", key, fullUrl);
            }
        } else {
            System.err.printf(
                "[%s] WARN: Confluence DR Doc URL field (customfield_12062) is empty.%n", key);
        }

        return new DRInfo(key, reviewers, creators, tinyUrl, fullUrl, pageId, pageTitle);
    }

    /**
     * Batch-fetches DR info for multiple tickets.
     * Errors on individual tickets are caught and logged; processing continues.
     *
     * @param ticketIds list of Jira ticket IDs
     * @return list of successfully resolved DRInfo records
     */
    public List<DRInfo> fetchAll(List<String> ticketIds) {
        List<DRInfo> results = new ArrayList<>();
        for (String ticketId : ticketIds) {
            try {
                results.add(fetch(ticketId));
            } catch (JiraDRException e) {
                System.err.printf("ERROR [%s]: %s%n", ticketId, e.getMessage());
            }
        }
        return results;
    }

    // ─── Private: Jira ─────────────────────────────────────────────────────────

    /**
     * Calls the Jira REST API via api.atlassian.com and returns the fields node.
     *
     * Correct URL: https://api.atlassian.com/ex/jira/{cloudId}/rest/api/3/issue/{key}
     * Wrong URL:   https://tdem.atlassian.net/rest/api/3/issue/{key}  ← returns 404
     */
    private JsonNode fetchJiraIssueFields(String ticketId) {
        String fields = String.join(",",
                FIELD_DR_REVIEWER, FIELD_DR_CREATOR, FIELD_CONFLUENCE_URL);
        String url = String.format(
                "%s/rest/api/3/issue/%s?fields=%s", jiraApiBase, ticketId, fields);

        System.out.printf("[%s] GET %s%n", ticketId, url);
        HttpResponse<String> response = doGet(url);

        switch (response.statusCode()) {
            case 200: break;
            case 401:
            case 403: throw new JiraAuthException("Jira issue " + ticketId);
            case 404: throw new JiraIssueNotFoundException(ticketId);
            default:
                throw new JiraDRException(String.format(
                    "[%s] Unexpected HTTP %d from Jira: %s",
                    ticketId, response.statusCode(), response.body()));
        }

        try {
            JsonNode root = mapper.readTree(response.body());
            JsonNode fieldsNode = root.path("fields");
            if (fieldsNode.isMissingNode())
                throw new JiraDRException(
                    "Jira response missing 'fields' node for: " + ticketId);
            return fieldsNode;
        } catch (IOException e) {
            throw new JiraDRException(
                "Failed to parse Jira response for: " + ticketId, e);
        }
    }

    // ─── Private: Confluence tiny URL resolution ───────────────────────────────

    /**
     * Resolves a Confluence /wiki/x/ short URL to the final page URL by
     * following the full redirect chain manually (HttpClient has redirects disabled).
     *
     * Confluence tiny URLs go through TWO hops:
     *   Hop 1: /wiki/x/AYD2QQ
     *       → /wiki/pages/tinyurl.action?urlIdentifier=AYD2QQ
     *   Hop 2: /wiki/pages/tinyurl.action?urlIdentifier=AYD2QQ
     *       → /wiki/spaces/DM/pages/1106673665/Page+Title   ← final URL with page ID
     *
     * The loop continues until it gets a non-redirect response or finds a URL
     * containing /pages/{id}/, with a safety cap to prevent infinite loops.
     */
    private String resolveConfluenceTinyUrl(String tinyUrl) {
        final int MAX_HOPS = 10;
        String currentUrl = tinyUrl;

        for (int hop = 1; hop <= MAX_HOPS; hop++) {
            System.out.printf("  [redirect hop %d] GET %s%n", hop, currentUrl);
            HttpResponse<String> response = doGet(currentUrl);
            int status = response.statusCode();

            if (status == 301 || status == 302 || status == 303
                    || status == 307 || status == 308) {

                String location = response.headers().firstValue("Location").orElse(null);
                if (location == null)
                    throw new JiraDRException(
                            "Redirect hop " + hop + " had no Location header for: " + currentUrl);

                // Make absolute if the server returned a relative path
                if (location.startsWith("/")) {
                    location = tenantUrl + location;
                }

                System.out.printf("  [redirect hop %d] → %s%n", hop, location);
                currentUrl = location;

                // Early exit: if we already have a /pages/{id}/ URL, no need to go further
                if (PAGE_ID_PATTERN.matcher(currentUrl).find()) {
                    System.out.printf("  [redirect hop %d] Found page ID in URL — stopping.%n", hop);
                    return currentUrl;
                }

            } else if (status == 200) {
                // Final destination reached without finding /pages/{id}/ in URL
                System.err.printf(
                        "WARN: Redirect chain ended at HTTP 200 without a /pages/{{id}}/ URL: %s%n",
                        currentUrl);
                return currentUrl;

            } else if (status == 401 || status == 403) {
                throw new JiraAuthException("Confluence redirect to: " + currentUrl);

            } else {
                throw new JiraDRException(String.format(
                        "Unexpected HTTP %d at redirect hop %d for URL: %s",
                        status, hop, currentUrl));
            }
        }

        throw new JiraDRException(
                "Exceeded maximum redirect hops (" + MAX_HOPS + ") for: " + tinyUrl);
    }

    /** Extracts the numeric page ID from a full Confluence /pages/{id}/ URL. */
    private String extractPageIdFromUrl(String url) {
        Matcher m = PAGE_ID_PATTERN.matcher(url);
        return m.find() ? m.group(1) : null;
    }

    // ─── Private: Confluence page title ───────────────────────────────────────

    /**
     * Calls the Confluence REST API to fetch the page title.
     *
     * URL: https://{tenant}.atlassian.net/wiki/rest/api/content/{pageId}
     */
    private String fetchConfluencePageTitle(String pageId, String ticketId) {
        String url = String.format(
                "%s/wiki/rest/api/content/%s", tenantUrl, pageId);

        HttpResponse<String> response = doGet(url);

        switch (response.statusCode()) {
            case 200: break;
            case 401:
            case 403: throw new JiraAuthException("Confluence page " + pageId);
            case 404: throw new ConfluencePageNotFoundException(pageId);
            default:
                throw new JiraDRException(String.format(
                    "[%s] Unexpected HTTP %d fetching Confluence page %s: %s",
                    ticketId, response.statusCode(), pageId, response.body()));
        }

        try {
            JsonNode root  = mapper.readTree(response.body());
            String   title = root.path("title").asText(null);
            if (title == null || title.isBlank())
                throw new JiraDRException(
                    "Confluence response missing 'title' for page ID: " + pageId);
            return title;
        } catch (IOException e) {
            throw new JiraDRException(
                "Failed to parse Confluence response for page ID: " + pageId, e);
        }
    }

    // ─── Private: Field extraction helpers ────────────────────────────────────

    /** Extracts displayNames from a multi-user or single-user Jira custom field. */
    private List<String> extractUserList(JsonNode fields, String fieldId) {
        JsonNode     node  = fields.path(fieldId);
        List<String> names = new ArrayList<>();
        if (node.isMissingNode() || node.isNull()) return names;

        if (node.isArray()) {
            node.forEach(user -> {
                String name = user.path("displayName").asText(null);
                if (name != null) names.add(name);
            });
        } else if (node.isObject()) {
            String name = node.path("displayName").asText(null);
            if (name != null) names.add(name);
        }
        return names;
    }

    /** Extracts a plain-string custom field. Returns null if absent or blank. */
    private String extractStringField(JsonNode fields, String fieldId) {
        JsonNode node = fields.path(fieldId);
        if (node.isMissingNode() || node.isNull()) return null;
        String value = node.asText(null);
        return (value == null || value.isBlank()) ? null : value;
    }

    // ─── Private: HTTP ─────────────────────────────────────────────────────────

    /** Executes an authenticated GET request using Atlassian Basic Auth. */
    private HttpResponse<String> doGet(String url) {
        String credentials = Base64.getEncoder()
                .encodeToString((email + ":" + apiToken).getBytes());

        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Basic " + credentials)
                .header("Accept", "application/json")
                .timeout(Duration.ofSeconds(15))
                .GET()
                .build();

        try {
            return httpClient.send(request, HttpResponse.BodyHandlers.ofString());
        } catch (IOException e) {
            throw new JiraDRException("Network error calling: " + url, e);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new JiraDRException("Request interrupted for: " + url, e);
        }
    }

    // ─── Private: Config helpers ───────────────────────────────────────────────

    /**
     * Loads config.properties from the classpath.
     * The file must be at: src/main/resources/config.properties
     */
    private static Properties loadProperties() {
        Properties props = new Properties();
        try (java.io.InputStream is = new java.io.FileInputStream("config.properties")) {
            if (is == null)
                throw new ConfigurationException(
                    CONFIG_FILE + " not found on classpath. " +
                    "Place it at: src/main/resources/" + CONFIG_FILE);

            props.load(is);
        } catch (IOException e) {
            throw new ConfigurationException(
                "Failed to read " + CONFIG_FILE + ": " + e.getMessage());
        }
        return props;
    }

    /** Reads a required property; throws ConfigurationException if missing/blank. */
    private static String requireProp(Properties props, String key) {
        String value = props.getProperty(key);
        if (value == null || value.isBlank())
            throw new ConfigurationException(
                "Missing or empty property '" + key + "' in " + CONFIG_FILE);
        return value.trim();
    }

    private static void requireNonBlank(String value, String name) {
        if (value == null || value.isBlank())
            throw new IllegalArgumentException("'" + name + "' must not be blank");
    }

    // ─── Main (quick smoke test) ───────────────────────────────────────────────

    public static void main(String[] args) {
        List<String> ticketIds = args.length > 0
                ? Arrays.asList(args)
                : List.of("DM-2199");

        JiraDRInfoFetcher fetcher;
        try {
            fetcher = JiraDRInfoFetcher.fromConfig();
        } catch (ConfigurationException e) {
            System.err.println("Configuration error: " + e.getMessage());
            System.exit(1);
            return;
        }

        List<DRInfo> results = fetcher.fetchAll(ticketIds);

        System.out.println("\n══════════════════════ RESULTS ══════════════════════");
        results.forEach(info -> System.out.println(info + "\n"));
        System.out.printf("Processed %d of %d ticket(s).%n",
                results.size(), ticketIds.size());
    }
}
