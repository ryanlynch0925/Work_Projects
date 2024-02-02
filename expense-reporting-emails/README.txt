# Expense Report Email Notifications

This Python script automates the process of sending email notifications to employees regarding outstanding expenses. It provides details about expenses that are not yet added to a report, in draft, or sent back for corrections.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Function Descriptions](#function-descriptions)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Overview

The script utilizes the `pandas` library and the Outlook API to interact with email functionality. It extracts information from the provided data, filters the top 40 records, and sends customized email notifications to employees.

## Prerequisites

Ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: `pandas`, `outlook`, etc. (install via `pip install -r requirements.txt`)
- An Outlook account with configured credentials (referenced in `config.py`)

## Usage

1. Clone the repository:
  git clone https://github.com/your-username/expense-report-emails.git
2. Navigate to the project directory:
  cd expense-report-notificatio
3.Install dependencies:
  pip install -r requirements.txt
4.Run the script:
  python expense_report_notifications.py

## Function Descriptions
create_email(outlook, unique_employees)
This function creates and sends customized email notifications to employees based on outstanding expenses. It includes details about corrections made to expense reports and the approval process involving managers.

### lean_and_filter(df)
This function cleans and filters the provided DataFrame to focus on records with outstanding expenses that have not been sent.

### create_email(outlook, unique_employees)
This function creates and sends email notifications to employees about corrections made to their expense reports. It supports both simple corrections and itemized details, informing employees and requesting manager approval.

## Contributing
Contributions are welcome! Please follow the guidelines outlined in CONTRIBUTING.md.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE.txt) file for details.
