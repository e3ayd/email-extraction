# Email Extraction and Categorization Script

This Python script extracts email addresses from an Excel file and categorizes them into two groups: Turkish emails and foreign emails. It uses a combination of domain-based detection and language detection through a Java library to identify Turkish emails.

## Features

- **Email Extraction**: Extracts all email addresses from an Excel file.
- **Language Detection**: Uses the `language-detection` Java library to determine if an email is Turkish based on the username and domain.
- **Domain-Based Categorization**: Categorizes emails with `.tr` domains as Turkish.
- **Excel Output**: Saves the categorized email addresses into a new Excel file with two separate sheets for Turkish and foreign emails.

## Prerequisites

- Python 3.x
- Java Development Kit (JDK)
- Pandas library (`pip install pandas`)
- JPype library (`pip install jpype1`)
- A `language-detection` Java library setup

## Setup

1. **Install Required Python Libraries**: Install the necessary Python libraries by running the following commands:
   ```bash
   pip install pandas jpype1
