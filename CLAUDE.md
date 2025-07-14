# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) project that synchronizes Gmail messages to Notion database pages. The integration runs in Google Sheets and provides a custom menu interface for configuration and manual execution.

## Core Architecture

### Main Components

1. **コード.js** - Main script file containing all functionality:
   - Settings management through Google Sheets
   - Gmail API integration for fetching labeled emails
   - Notion API integration for creating database pages
   - Trigger management for automated synchronization
   - Custom UI with Google Sheets menu

2. **無題.html** - Settings dialog interface:
   - HTML form for configuring Notion integration keys
   - Client-side validation and Google Apps Script integration
   - Provides user-friendly interface for entering sensitive configuration

3. **appsscript.json** - Project configuration:
   - Runtime: V8
   - Timezone: Asia/Tokyo
   - Exception logging: STACKDRIVER

### Key Functions

- `onOpen()` - Creates custom menu when spreadsheet opens
- `syncEmailsToNotion()` - Main synchronization function
- `createNotionPage()` - Handles Notion API page creation
- `getSettings()`/`saveSettings()` - Configuration management
- `createSettingsSheet()` - Sets up configuration spreadsheet

### Configuration System

The application uses a dedicated Google Sheets tab called "設定シート" (Settings Sheet) to store:
- Notion Integration Token
- Notion Database ID  
- Title property name for Notion database
- Gmail label to monitor
- Last processed timestamp for incremental sync

### Gmail Integration

- Monitors emails with specific Gmail label (default: "Notion")
- Performs incremental synchronization using last processed timestamp
- Processes up to 50 emails per execution to avoid timeout

### Notion Integration

- Uses Notion API v2022-06-28
- Creates structured pages with email metadata and content
- Formats email content as rich text blocks with headers and dividers
- Truncates email body to 2000 characters for API compliance

## Development Notes

### Google Apps Script Environment

- No traditional build/test/lint commands - development happens in Google Apps Script web editor
- Runtime environment is V8 with Google Apps Script APIs
- No package.json or traditional dependency management

### Deployment

1. Code is deployed through Google Apps Script web interface
2. No CI/CD pipeline - manual deployment only
3. Triggers are configured through the application's custom menu

### Testing

- Test through the "手動で同期を実行" (Manual Sync) menu option
- Monitor execution logs through Google Apps Script dashboard
- Use stackdriver logging for error tracking

### Key Limitations

- Google Apps Script execution time limits (6 minutes for triggers, 30 minutes for manual execution)
- Notion API rate limits and payload size restrictions
- Gmail API quotas and batch processing limitations

### Error Handling

- Comprehensive try-catch blocks with Japanese error messages
- User-friendly alerts through Google Sheets UI
- Console logging for debugging and monitoring

### Security Considerations

- Notion integration tokens stored in spreadsheet (consider using PropertiesService for production)
- HTML form uses password input type for sensitive keys
- No authentication beyond Google account access to spreadsheet