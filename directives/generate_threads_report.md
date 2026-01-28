# Generate Threads Daily Report Directive

This directive outlines the process to fetch posts from a Threads user, analyze them using Gemini, and send an email report.

## Goal
Fetch the last 24 hours of posts from a specific Threads user (@choi.openai), summarize them, extract actionable insights, and email the report.

## Inputs
- **Target User**: `choi.openai` (hardcoded for now, or passed as arg)
- **Time Window**: Last 24 hours

## Tools & Scripts
1.  **Fetch Data**: `execution/fetch_threads_posts.py`
    -   **Method**: Playwright (Headless Browser) due to API limitations.
    -   **Input**: Target username
    -   **Output**: `.tmp/threads_posts.json` (List of posts with `timestamp` and `text`)
2.  **Analyze & Send**: `execution/analyze_and_report.py`
    -   **Input**: `.tmp/threads_posts.json`
    -   **Process**:
        -   Filter posts by timestamp (now - 24h).
        -   Format text for LLM.
        -   Call Gemini API for summary and insights.
        -   Auth with Gmail API.
        -   Send email to recipient in `.env`.
    -   **Output**: Console log of success/failure.

## Workflow Steps
1.  Run `python execution/fetch_threads_posts.py`
2.  Run `python execution/analyze_and_report.py`

## Error Handling
-   If `fetch_threads_posts.py` fails (e.g., API change, login required), log error and exit.
-   If no new posts found in 24h, `analyze_and_report.py` should log "No new posts" and skip email (or send "No updates" email if preferred).
