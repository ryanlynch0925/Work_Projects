# Tidal Wave Auto Spa Fleet Invoice Sender

The Tidal Wave Auto Spa Fleet Invoice Sender is a Python script crafted to automate the process of sending monthly fleet invoices to clients. The script extracts relevant information from PDF invoices, categorizes accounts based on due amounts, and sends tailored emails with attached invoices.

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

The Tidal Wave Auto Spa Fleet Invoice Sender streamlines the communication process with fleet clients by automatically generating emails containing the respective monthly invoice. The script categorizes accounts based on due amounts and customizes the email content accordingly.

## Features

- Extracts past-due information from PDF invoices.
- Sends personalized emails to clients based on their account status.
- Supports payment options such as check and credit card.
- Ensures timely communication and improves payment processing efficiency.

## Getting Started

### Prerequisites

Before using the Tidal Wave Auto Spa Fleet Invoice Sender, ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: PyPDF2, win32com.client (install via pip install -r requirements.txt)

### Installation

1. Clone the repository:
   git clone https://github.com/your-username/Fleet-Invoice-Emailer.git

2. Navigate to the project directory:
    cd Fleet-Invoice-Emailer

3. Install dependencies:
    pip install -r requirements.txt

### Usage

1.Set up your Outlook credentials and other configurations in the config.py file.
2.Prepare your expense report data in an Excel sheet.
3.Run the script:
    python Fleet-Invoice-Emailer.py
4.Follow the prompts to initiate the corrections communication process.

### Configuration

Adjust the 'config.py' file to customize the email templates, Outlook settings, and other configurations.

### License
This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt)file for details.
