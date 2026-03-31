# Confluence Comment Section Extractor

## Overview

`ConfluenceCommentSectionExtractor` reads a batch input Excel file, processes each Confluence `pageId`, fetches inline and footer comments through the Confluence REST API, maps inline comments to the nearest page heading/section, and writes all results into a **single consolidated Excel workbook**.

This version is designed for migration/review workflows where each input row contains business metadata such as module and function information.

## What the program does

For each row in the input workbook, the program:

1. Reads the `pageId` from the first worksheet.
2. Fetches the page body in Confluence storage format.
3. Parses headings (`h1`-`h6`) and inline comment markers.
4. Resolves each inline comment to the nearest preceding section heading.
5. Fetches inline comments and footer/page-level comments.
6. Writes all extracted comments into one output workbook.
7. Prints a per-page summary to the console.

## Current source files

Primary implementation:

- `src/main/java/ConfluenceCommentSectionExtractor.java`

Jira DR report utility:

- `src/main/java/JiraDRReportGenerator.java`

Shared Jira URL resolution helper:

- `src/main/java/JiraDRInfoFetcher.java`

Legacy single-page extractor still present in the repo:

- `src/main/java/ConfluenceCommentsExtractor.java`

This README covers both the **section-aware Confluence comment extractor** and the **Jira DR report generator**.

## Prerequisites

- Java 11+
- Maven 3.6+ (recommended)
- Access to the target Confluence site
- Atlassian email + API token
- An input `.xlsx` file in the required format
- A writable output folder at `D:\outputs\` (the current code writes there)

## Configuration

The application reads credentials and base URL from `config.properties` in the project root.

Example:

```properties
base.url=https://your-domain.atlassian.net/wiki
tenant.url=https://your-domain.atlassian.net/wiki
email=your.name@example.com
api.token=your_api_token_here
```

### Important

- Do **not** commit real API tokens to source control.
- `ConfluenceCommentSectionExtractor` uses `base.url`, `email`, and `api.token`.
- `JiraDRReportGenerator` uses `tenant.url`, `email`, and `api.token`.
- If `config.properties` is missing or invalid, the program fails during initialization.

## Input Excel format

The program reads:

- the **first sheet only**
- starting from **row 2** (row 1 is treated as header)
- columns by **position**, not by header name text
- for accurate per-row summary counts, keep `pageId` unique per input row (current summary counting is keyed by `pageId`)

### Required column order

| Column | Header text (recommended) | Required | Notes |
|---|---|---:|---|
| A | Module | No | Copied to output as-is |
| B | Legacy ID | No | Copied to output as-is |
| C | New ID | No | Copied to output as-is |
| D | Function Name | No | Copied to output as-is |
| E | Function Type | No | Copied to output as-is |
| F | pageId | Yes | Row is processed only when this value is non-blank |
| G | Reviewer Name | No | Supported by current code and copied to output |

### Example input sheet

| Module | Legacy ID | New ID | Function Name | Function Type | pageId | Reviewer Name |
|---|---|---|---|---|---|---|
| Billing | LEG-101 | NEW-101 | Create Invoice | API | 123456789 | Alice |
| Claims | LEG-205 | NEW-205 | Submit Claim | Screen | 987654321 | Bob |

### Notes about input parsing

- Header names are for readability only; the code uses fixed column indexes.
- Numeric `pageId` values are converted to whole-number strings.
- Rows with blank `pageId` are skipped.

## Output Excel format

The program writes a single workbook to:

`D:\outputs\all_confluence_comments_yyyyMMdd_HHmmss.xlsx`

Sheet name:

- `Comments`

### Output columns

| Column | Description |
|---|---|
| Module | From input sheet |
| Legacy ID | From input sheet |
| New ID | From input sheet |
| Function Name | From input sheet |
| Function Type | From input sheet |
| Type | `Inline` or `Footer` |
| Section | Nearest resolved heading for inline comments, or `Page Level` for footer comments |
| Highlighted Text | Inline comment selection text |
| Comment ID | Confluence comment ID |
| Author | Comment creator display name |
| Created Date | Comment creation timestamp |
| Last Updated | Last update timestamp |
| Status | Usually `open` or `resolved` |
| Comment Text | Plain-text comment body |
| Link | Clickable URL back to the comment in Confluence |
| Reviewer Name | From input sheet |

### Output formatting features

- bold dark-blue header row
- alternating row background styling
- wrapped text for readability
- frozen header row
- auto-filter on the full data range
- clickable hyperlink in the `Link` column

## Special output behavior

The current code contains merge behavior for comments where `Highlighted Text` (`originalSelection`) is blank:

- no new Excel row is created for that comment
- the comment text is appended to the **previous row's** `Comment Text` cell
- the appended text is added on a new line

Because of that behavior, the summary count is intended to reflect the number of **rows actually written to the output workbook** for each input `pageId`.

## Console summary

At the end of execution, the program prints a per-page summary with these fields:

- `Module`
- `Legacy ID`
- `New ID`
- `Function Name`
- `Function Type`
- `pageId`
- `Number of comments`

The default output is a tab-separated summary (`printExecutionSummary`) that can be copied into Excel.

The code also contains a `FlipTable`-based pretty table method (`printExecutionSummaryTabular`), but the current `main` method calls the tab-separated summary printer.

## Dependencies used by the current implementation

- Apache POI (`poi`, `poi-ooxml`) for Excel I/O and formatting
- `org.json` for JSON parsing
- `jsoup` for HTML/XML parsing and HTML stripping
- `fliptables` for optional tabular console summary output

## Jira DR Report Generator

`JiraDRReportGenerator` converts a Jira Excel export into a cleaned DR mapping workbook and resolves Confluence `Page ID` values from DR document URLs (including short/tiny URLs through authenticated redirect hops).

### Jira input Excel format

The generator expects:

- workbook sheet name: `Your Jira Issues`
- header row on row 1
- these input columns by header text:
  - `Key`
  - `Summary`
  - `Legacy Function ID`
  - `New Function ID`
  - `Confluence DR Doc`
  - `DR_Reviewer`

### Jira output Excel format

Output file name pattern:

- `output_yyyyMMddHHmmss.xlsx`

Output sheet name:

- `DR Report`

Output columns:

1. `Module`
2. `Legacy ID`
3. `New ID`
4. `Function Name`
5. `Function Type`
6. `Page ID`
7. `Reviewer Name`
8. `Confluence URL`
9. `Full URL`
10. `Jira Ticket ID`

### Jira field derivation rules

- `Module`: characters at index `1..3` of `New Function ID` (example: `WCMN00190` -> `CMN`)
- `Function Type` from first character of `New Function ID`:
  - `W` -> `Screen`
  - `B` -> `Batch`
  - `R` -> `Report`
  - `I` -> `Interface`
- `Page ID`: extracted directly from Confluence URL when possible, otherwise resolved from redirect chain (`/wiki/x/...`, `tinyurl.action`, `viewpage.action`, `resumedraft.action`)

### Run Jira DR Report Generator

Build first:

```powershell
Set-Location "D:\ConfluenceCrawler_java\confluence_crawler"
mvn clean package
```

Run from compiled classes (works regardless of JAR main class):

```powershell
Set-Location "D:\ConfluenceCrawler_java\confluence_crawler"
java -cp target\classes JiraDRReportGenerator "D:\path\to\jira-search-result.xlsx"
```

Run with explicit output directory:

```powershell
Set-Location "D:\ConfluenceCrawler_java\confluence_crawler"
java -cp target\classes JiraDRReportGenerator "D:\path\to\jira-search-result.xlsx" "D:\outputs"
```

## Runtime usage

The program expects exactly one command-line argument:

```text
java -jar target\comments-extractor-1.0.0.jar <inputExcelFile>
```

Example:

```powershell
java -jar target\comments-extractor-1.0.0.jar "D:\inputs\confluence_pages.xlsx"
```

## How section resolution works

Inline comments are matched to sections by:

1. fetching the page body in storage format
2. walking the document in order
3. tracking the most recent heading (`h1`-`h6`)
4. mapping each `<ac:inline-comment-marker ac:ref="...">` to that heading

If an inline comment cannot be matched, the fallback section text is:

- `(Section not resolved)`

If a marker appears before any heading, it is grouped under:

- `(Before first heading)`

Footer comments are written as:

- `Type = Footer`
- `Section = Page Level`

## Troubleshooting

### Authentication errors

- verify `email` and `api.token` in `config.properties`
- confirm the token is active
- make sure the base URL includes `/wiki` if required by your Confluence Cloud URL

### 404 / page not found

- confirm the `pageId` is correct
- verify the account has permission to access the page

### Output folder issues

The current code writes to `D:\outputs\...`.

Create the folder first if it does not exist:

```powershell
New-Item -ItemType Directory -Force -Path "D:\outputs"
```

### No rows written for a page

- check whether `pageId` is blank in the input row
- confirm the page actually has comments
- verify the account can read inline and footer comments

## Security recommendation

Do not store real credentials in the repository. Prefer one of these approaches:

- keep `config.properties` local and uncommitted
- use a separate environment-specific config file
- rotate API tokens regularly

## Reference

- Confluence REST API: <https://developer.atlassian.com/cloud/confluence/rest/>
- Apache POI: <https://poi.apache.org/>
