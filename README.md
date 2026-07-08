# PDF Payment Parser & Matching System

Business process automation built with Google Apps Script to extract structured payment information from PDF receipts, validate records, and match them against payment data stored in Google Sheets.

## Overview

This project automates a manual payment verification workflow by processing PDF receipts stored in Google Drive. It extracts relevant payment information, transforms unstructured text into structured records, and validates the results against an internal payment dataset.

The automation was designed to reduce repetitive manual work, improve data consistency, and support large-scale document processing.

## Features

- Batch PDF processing
- PDF text extraction using Google Docs
- Structured data parsing with regular expressions
- Multi-field record matching
- Data normalization
- Progress tracking
- Resume processing after execution timeout
- Google Sheets integration
- Google Drive integration

## Workflow

```text
Google Drive (PDF)
        │
        ▼
Extract Text
        │
        ▼
Parse Structured Data
        │
        ▼
Normalize Values
        │
        ▼
Match Against Records
        │
        ▼
Generate Output
        │
        ▼
Google Sheets
```

## Technology Stack

| Category | Technology |
|----------|------------|
| Language | Google Apps Script (JavaScript) |
| Spreadsheet | Google Sheets |
| Storage | Google Drive |
| Document Processing | Google Docs API |
| Parsing | Regular Expressions |
| Automation | Time-Based Triggers |
| State Management | Script Properties |

## Technical Highlights

### Batch Processing

Processes PDF files in configurable batches to work within Google Apps Script execution limits.

### PDF Processing

Converts PDF files into temporary Google Documents, extracts the text, and automatically removes temporary files after processing.

### Parsing Engine

Extracts the following information from payment receipts:

- Payment Amount
- Currency
- Campaign
- KOL
- Statement of Work (SOW)
- Bank Name
- Bank Account Number

### Matching Engine

Validates records using multiple fields:

- Campaign
- KOL
- Amount
- Bank Name
- Account Number
- SOW

Text normalization is applied before comparison to improve matching accuracy.

### Resume Mechanism

Automatically resumes processing from the last completed record after execution timeout.

## Repository Structure

```text
.
├── Code.gs
├── pdf_ui.html
├── README.md
└── appsscript.json
```

## Skills Demonstrated

- Business Process Automation
- Document Processing
- PDF Parsing
- Regular Expressions
- Workflow Automation
- Batch Processing
- Data Validation
- Google Workspace Automation

## License

This project is licensed under the MIT License.
