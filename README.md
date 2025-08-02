# Mail Merge Add-on for Google Docs

This add-on streamlines the process of creating personalized letters and documents by merging data from Google Sheets into Google Docs. It supports both individual and batch document generation, with options for PDF export and custom number formatting.

## Features

- Select Google Docs and Sheets from Google Drive
- Choose merge fields and customize formatting
- Merge all records or select a specific range
- Export merged documents as Google Docs or PDFs
- Consistent error handling and user-friendly interface

## Installation

1. Clone or download this repository.
2. Open the project in Google Apps Script.
3. Deploy as an add-on or web app within your Google Workspace domain.

## Usage

1. Launch the add-on from Google Docs.
2. Select the document type and template.
3. Pick a Google Sheet and choose merge fields.
4. Configure formatting and merge options.
5. Generate merged documents or PDFs.

## File Structure

- `helpersFunctions.js`: Utility functions for formatting and merging.
- `WebApp.js`: Server-side logic and API connectors.
- `JS.html`: Client-side UI and error handling.
- `loadExistingList.js`, `mergeLetters.js`, `MainPage.js`: UI and workflow logic.

## Troubleshooting

- Ensure merge fields match spreadsheet headers exactly.
- Errors are displayed in the web interface for easy debugging.

## License

MIT

