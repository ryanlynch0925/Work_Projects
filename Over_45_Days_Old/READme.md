# Expense Report Summary Generator

The Expense Report Summary Generator is a Python script designed to process and summarize expense reports. It identifies employees with expenses over 45 days old, generates a summary for each of them, and saves the summary to an Excel file.

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

The Expense Report Summary Generator is a Python script that automates the identification and summarization of expense reports for employees with expenses over 45 days old.

## Features

- Identifies employees with expenses over 45 days old.
- Generates a summary for each identified employee, including total expenses and other relevant details.
- Saves the generated summaries to an Excel file.

## Getting Started

### Prerequisites

Before using the Expense Report Summary Generator, ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: `pandas`, `openpyxl`, etc. (install via `pip install -r requirements.txt`)

### Installation

1. Clone the repository:
   git clone https://github.com/your-username/over-45-days-old-report.git

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