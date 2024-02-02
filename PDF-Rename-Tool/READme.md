# PDF Invoice Organizator

The PDF Invoice Organizer is a Python script designed to organize PDF invoices by extracting relevant information such as amount, invoice number, and date from the contents. It then renames and moves the files to a new location based on this information.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Contributing](#contributing)
- [License](#license)

## Introduction

The PDF Invoice Organizer automates the process of organizing PDF invoices by extracting key information and renaming the files accordingly. It helps maintain a structured and easily accessible invoice filing system.

## Features

- Automated Organization: Extracts essential details from PDF invoices.
- File Renaming: Renames PDF files based on extracted information for clarity.
- Scheduled Deletion: Deletes original PDFs after organization to reduce clutter.

## Getting Started

### Prerequisites

Before using the PDFForm Automator, ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: `pandas`, `openpyxl`, `pywin32`, `docx` (install via `pip install -r requirements.txt`)

### Installation

1. Clone the repository:
   git clone https://github.com/your-username/pdf-form-automator.git

2. Navigate to the project directory:
    cd expense-report-corrections

3. Install dependencies:
    pip install -r requirements.txt

### Usage

1.Set up your Outlook credentials and other configurations in the config.py file.
2.Prepare your expense report data in an Excel sheet.
3.Run the script:
    python expense_report_corrections.py
4.Follow the prompts to initiate the corrections communication process.

### Configuration

Adjust the 'config.py' file to customize the email templates, Outlook settings, and other configurations.

### License
This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt)file for details.