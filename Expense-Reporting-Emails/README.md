# Expense Report Corrections Tool

This tool automates the process of notifying employees about corrections made to their expense reports and sending necessary information to managers for approval.

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

The Expense Report Corrections Tool is a Python script that utilizes the `pandas` library and the Outlook API to automate the communication process regarding corrections made to employee expense reports.

## Features

- Sends automated emails to employees regarding corrections in their expense reports.
- Notifies managers for review and approval after corrections are made.
- Supports itemized expenses with detailed information.
- Configurable email templates and signature.

## Getting Started

### Prerequisites

Ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: `pandas`, `outlook`, etc. (install via `pip install -r requirements.txt`)

### Installation

1. Clone the repository:
   git clone https://github.com/your-username/Expense-Reporting-Emails.git

2. Navigate to the project directory:
    cd Expense-Reporting-Emails

3. Install dependencies:
    pip install -r requirements.txt

### Usage

1.Set up your Outlook credentials and other configurations in the config.py file.
2.Prepare your expense report data in an Excel sheet.
3.Run the script:
    python fixedreports.py
4.Follow the prompts to initiate the corrections communication process.

### Configuration

Adjust the 'config.py' file to customize the email templates, Outlook settings, and other configurations.

### License
This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt)file for details.
