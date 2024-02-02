# PDFForm Automator

The PDFForm Automator is a Python script designed to automate the process of generating PDFs from an Excel file and saving them with the respective site names.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Introduction

The PDFForm Automator automates the steps of gathering data, inputting it into an Excel form, converting it to PDF, and saving the files with the corresponding site names.

## Features

- Generates PDFs from Excel data.
- Saves PDFs with the respective site names.
- Handles errors and provides a detailed report.

## Getting Started

### Prerequisites

Before using the PDFForm Automator, ensure you have the following installed:

- Python (3.x recommended)
- Required Python packages: `pandas`, `win32com.client`, `docx` (install via `pip install -r requirements.txt`)

### Installation

1. Clone the repository:
   git clone https://github.com/your-username/Utilities-Form-Creator.git

2. Navigate to the project directory:
    cd Utilities-Form-Creator

3. Install dependencies:
    pip install -r requirements.txt

### Usage

1.Set up your Outlook credentials and other configurations in the config.py file.
2.Prepare your expense report data in an Excel sheet.
3.Run the script:
    python Utilities-Form-Creator.py
4.Follow the prompts to initiate the corrections communication process.

### License
This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt)file for details.
