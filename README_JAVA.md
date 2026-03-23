# Confluence Comments to Excel Extractor (Java)

## Overview
Java application to extract all comments from a Confluence page and export them to a formatted Excel file.

## Prerequisites
- Java 11 or higher
- Maven 3.6+ or Gradle 7.0+ (for dependency management)
- Atlassian API token
- Confluence page ID

## Project Structure
```
confluence-comments-extractor/
├── ConfluenceCommentsExtractor.java
├── pom.xml (Maven)
├── build.gradle (Gradle - alternative)
└── README.md
```

## Setup Instructions

### 1. Get Your Atlassian API Token
1. Visit https://id.atlassian.com/manage-profile/security/api-tokens
2. Click "Create API token"
3. Name it (e.g., "Confluence Comments Extractor")
4. Copy and save the token securely

### 2. Find Your Confluence Page ID

**Method A: From URL**
- Open your Confluence page
- URL format: `https://tdem.atlassian.net/wiki/spaces/SPACE/pages/PAGE_ID/Page+Title`
- Extract the number after `/pages/`

**Method B: Page Information**
- Open the page → Click "..." menu → "Page Information"
- Page ID is in the URL

**Method C: Edit Mode**
- Edit the page and check URL for `pageId=` parameter

### 3. Configure the Application

Edit `ConfluenceCommentsExtractor.java` and update these constants:

```java
private static final String BASE_URL = "https://tdem.atlassian.net/wiki";
private static final String PAGE_ID = "123456789";  // Your page ID
private static final String EMAIL = "your.email@example.com";
private static final String API_TOKEN = "your_api_token_here";
```

### 4. Build and Run

#### Option A: Using Maven

**Build:**
```bash
mvn clean package
```

**Run:**
```bash
java -jar target/comments-extractor-1.0.0.jar
```

Or run directly:
```bash
mvn exec:java -Dexec.mainClass="ConfluenceCommentsExtractor"
```

#### Option B: Using Gradle

**Build:**
```bash
gradle shadowJar
```

**Run:**
```bash
java -jar build/libs/confluence-comments-extractor-1.0.0.jar
```

Or run directly:
```bash
gradle run
```

#### Option C: Compile and Run Manually

If you don't have Maven/Gradle, download dependencies manually:

1. Download required JARs:
   - poi-5.2.5.jar
   - poi-ooxml-5.2.5.jar
   - xmlbeans-5.1.1.jar
   - commons-compress-1.24.0.jar
   - commons-collections4-4.4.jar
   - json-20231013.jar

2. Compile:
```bash
javac -cp "lib/*" ConfluenceCommentsExtractor.java
```

3. Run:
```bash
java -cp ".:lib/*" ConfluenceCommentsExtractor
```

## Output

The application creates an Excel file: `confluence_comments_YYYYMMDD_HHMMSS.xlsx`

### Excel Columns:
- **Comment ID**: Unique identifier
- **Author**: Display name of commenter
- **Created Date**: Initial creation timestamp
- **Last Updated**: Last modification timestamp
- **Comment Text**: Comment content (HTML stripped)

### Excel Features:
- Blue header with white text
- Auto-filter enabled on all columns
- Text wrapping for readability
- Optimized column widths

## Dependencies

### Apache POI (Excel Processing)
```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.5</version>
</dependency>
```

### JSON Processing
```xml
<dependency>
    <groupId>org.json</groupId>
    <artifactId>json</artifactId>
    <version>20231013</version>
</dependency>
```

## Troubleshooting

### Authentication Issues (HTTP 401)
- Verify email and API token are correct
- Check API token hasn't expired
- Ensure no extra spaces in credentials

### Page Not Found (HTTP 404)
- Confirm page ID is correct
- Verify you have view permissions for the page
- Check if page is in a restricted space

### Build Errors
**Maven:**
```bash
mvn clean install -U
```

**Gradle:**
```bash
gradle clean build --refresh-dependencies
```

### Runtime Errors
- Ensure Java 11+ is installed: `java -version`
- Verify all dependencies are downloaded
- Check network connectivity to Atlassian

## Advanced Usage

### Extract from Multiple Pages

Modify the `main` method:

```java
public static void main(String[] args) {
    String[] pageIds = {"123456", "789012", "345678"};
    List<Comment> allComments = new ArrayList<>();
    
    for (String pageId : pageIds) {
        List<Comment> comments = extractComments(BASE_URL, pageId, EMAIL, API_TOKEN);
        if (comments != null) {
            allComments.addAll(comments);
        }
    }
    
    String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
    createExcelReport(allComments, "all_comments_" + timestamp + ".xlsx");
}
```

### Use Environment Variables for Security

```java
private static final String EMAIL = System.getenv("CONFLUENCE_EMAIL");
private static final String API_TOKEN = System.getenv("CONFLUENCE_API_TOKEN");
```

Set environment variables:
```bash
# Linux/Mac
export CONFLUENCE_EMAIL="your.email@example.com"
export CONFLUENCE_API_TOKEN="your_token_here"

# Windows
set CONFLUENCE_EMAIL=your.email@example.com
set CONFLUENCE_API_TOKEN=your_token_here
```

### Filter Comments by Date

Add this method:

```java
private static List<Comment> filterByDate(List<Comment> comments, LocalDateTime cutoffDate) {
    return comments.stream()
        .filter(c -> {
            try {
                LocalDateTime created = LocalDateTime.parse(
                    c.createdDate.substring(0, 19),
                    DateTimeFormatter.ISO_LOCAL_DATE_TIME
                );
                return created.isAfter(cutoffDate);
            } catch (Exception e) {
                return true;
            }
        })
        .collect(Collectors.toList());
}
```

### Command Line Arguments

Enhance the main method to accept arguments:

```java
public static void main(String[] args) {
    if (args.length < 3) {
        System.out.println("Usage: java -jar app.jar <page_id> <email> <api_token>");
        System.exit(1);
    }
    
    String pageId = args[0];
    String email = args[1];
    String apiToken = args[2];
    
    // Rest of the code...
}
```

## Integration with Spring Boot

For enterprise use, integrate with Spring Boot:

1. Add Spring dependencies
2. Create a REST controller
3. Use `@ConfigurationProperties` for settings
4. Implement async processing with `@Async`

## Performance Considerations

- **Large comment sets**: The API uses pagination (100 comments per request)
- **Rate limiting**: Atlassian APIs have rate limits (check your plan)
- **Memory**: Large HTML comments are cleaned in-memory
- **Connection timeout**: Default is system timeout; customize if needed

## Security Best Practices

1. **Never commit credentials** to version control
2. Use environment variables or external config files
3. Rotate API tokens regularly
4. Use read-only tokens when possible
5. Log security events appropriately

## License

This is a utility tool for internal use. Ensure compliance with your organization's policies and Atlassian's API terms of service.

## Support

For issues or questions:
- Check Atlassian REST API documentation: https://developer.atlassian.com/cloud/confluence/rest/
- Review Apache POI documentation: https://poi.apache.org/

## Version History

- **1.0.0** (2025-01-30)
  - Initial release
  - Basic comment extraction
  - Excel export with formatting
  - HTML content cleaning
