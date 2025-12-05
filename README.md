# 🧾 AI Payment Slip Parser

AI-powered application for extracting structured data from bank payment
slips and exporting them into a clean Excel file.

------------------------------------------------------------------------

# 📑 Table of Contents

1.  Introduction
2.  Key Features
3.  System Requirements
4.  Installation
5.  User Guide
    -   1.  Main Interface

    -   2.  API Key Setup

    -   3.  Select Source Folder

    -   4.  Settings

    -   5.  Progress & Logs

    -   6.  Start / New / Exit Buttons
6.  Output File Structure
7.  Troubleshooting
8.  Notes

------------------------------------------------------------------------

# Introduction

AI Payment Slip Parser is a graphical Python application that reads bank
payment documents (PDFs or images), extracts structured fields using AI,
and saves all results into an Excel spreadsheet.

The application is designed for users who frequently categorize, record,
or analyze payment slips and want a fast, automated workflow.

------------------------------------------------------------------------

# Key Features

-   AI-powered extraction of payment slip fields\
-   Clean and intuitive GUI\
-   Batch processing of multiple files\
-   Dynamic extraction mode for capturing additional non-standard
    fields\
-   Live progress bar and event logs\
-   Configurable output folder\
-   Requires only a valid Gemini API key

------------------------------------------------------------------------

# System Requirements

-   Windows 10 or 11\
-   Python 3.10 or newer\
-   Dependencies from `requirements.txt`

------------------------------------------------------------------------

# Installation

1.  Install Python (3.10+).\
2.  Install dependencies:

```{=html}
<!-- -->
```
    pip install -r requirements.txt

3.  Obtain a valid Google Gemini API key.\
4.  Place your API key inside **api.txt** in the project root folder.

------------------------------------------------------------------------

# User Guide

## 1. Main Interface

![Main GUI](screenshots/1_gui.png)

## 2. API Key Setup

Insert your Gemini API key into `api.txt`.

![API Setup](screenshots/2_api.png)

## 3. Select Source Folder

Choose the folder containing the payment slips you want to process.

![Source Folder](screenshots/3_source_folder.png)

## 4. Settings

Configure extraction mode and output location.

![Settings](screenshots/4_settings.png)

## 5. Progress & Logs

Monitor progress and status messages.

![Progress & Logs](screenshots/5_progress_logs.png)

## 6. Start / New / Exit Buttons

Start processing, reset, or close the application.

![Start, New, Exit](screenshots/6_start_new_exit.png)

------------------------------------------------------------------------

# Output File Structure

The final Excel file contains the following fields:

-   `charging_bank`\
-   `transaction_id`\
-   `transaction_date`\
-   `amount`\
-   `fees`\
-   `iban_from`\
-   `iban_to`\
-   `credit_bank`\
-   `account_holder_same_bank`\
-   `file_name`\
-   `bank_name_header`

All data is saved in the output folder selected by the user.

------------------------------------------------------------------------

# Troubleshooting

### "API key not found"

Verify that `api.txt` exists and contains a valid API key.

### "No files found"

Ensure the selected source folder contains supported file types.

### Extraction errors

Check the log panel for detailed error messages.

------------------------------------------------------------------------

# Notes

-   The API key stays locally on your machine and is ignored by Git.\
-   Only text content is transmitted to the AI model.
