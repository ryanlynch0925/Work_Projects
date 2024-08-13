# Expense Deduction Notifier

The Expense Deduction Notifier is a Python script designed to automate the process of notifying employees about personal expense deductions and sending necessary information for their payroll deductions.

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

The Expense Deduction Notifier reads personal expense data from an Excel file, calculates the total reimbursement, and sends personalized summary emails to each employee. The script assists in notifying employees about their personal charge balance and informs them about the upcoming payroll deductions.

## Features

- Sends personalized summary emails to employees regarding personal expense deductions.
- Provides details about the total personal charge balance deducted from the next payroll.
- Ensures effective communication and transparency in the expense reimbursement process.

## Getting Started

### Prerequisites

Before using the Expense Deduction Notifier, ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: `pandas`, `win32com.client` (install via `pip install -r requirements.txt`)

### Installation

1. Clone the repository:
   git clone https://github.com/your-username/expense-deduction-notifier.git

2. Navigate to the project directory:
    cd expense-deduction-notifier

3. Install dependencies:
    pip install -r requirements.txt

### Usage

1.Set up your Outlook credentials and other configurations in the config.py file.
2.Prepare your expense report data in an Excel sheet.
3.Run the script:
    python expense-deduction-notifier.py
4.Follow the prompts to initiate the corrections communication process.

### Configuration

Adjust the 'config.py' file to customize the email templates, Outlook settings, and other configurations.

### License
This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt)file for details.