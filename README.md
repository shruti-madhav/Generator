# PO Document Generator

This Python program automates the process of generating Word documents based on data extracted from Excel files. Specifically, it creates a new `.docx` file for each Purchase Order (PO) number found in the Excel file.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Requirements](#requirements)
- [Code Explanation](#code-explanation)
- [Excel File Structure](#excel-file-structure)
- [Template Document Explanation](#template-document-explanation)
- [Screenshots](#screenshots)
- [Contributing](#contributing)
- [License](#license)

## Introduction

The PO Document Generator is designed to streamline the creation of Word documents by automatically processing Excel files and generating a new document for each PO number. This repository contains all the necessary code and instructions to set up and run the program.

## Features

- Extracts data from Excel files.
- Automatically generates a `.docx` file for each PO number.
- Easy to set up and run.

## Installation

1. Clone this repository to your local machine:

    ```bash
    git clone https://github.com/yourusername/po-document-generator.git
    ```

2. Navigate to the project directory:

    ```bash
    cd po-document-generator
    ```

3. Create a virtual environment (optional but recommended):

    ```bash
    python3 -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

4. Install the required Python packages:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Place your Excel file in the project directory.

2. Update the `main.py` file if necessary to match your Excel file's structure.

3. Run the program:

    ```bash
    python main.py
    ```

4. The generated `.docx` files will be saved in the output directory.

## Requirements

- Python 3.10.0
- The necessary Python packages are listed in the `requirements.txt` file.

## Code Explanation

The `main.py` file contains the logic for reading the Excel file, processing the data, and generating the Word documents. Here's a brief overview of the key components:

- **Reading Excel Data**: The program reads the data from the Excel file using `pandas`. It then iterates through each row, extracting relevant information such as PO number, product details, and other necessary fields.

- **Template Handling**: The program opens a pre-defined Word document template using `python-docx`. Placeholders in the template are identified and replaced with actual values from the Excel data.

- **Product Rows**: The program dynamically adds rows to the product table in the template for each item listed under the PO number.

- **Saving Output**: Finally, the customized document is saved with a unique name corresponding to the PO number.

## Excel File Structure

The Excel file should contain the following columns:

- **PO Number**: The unique identifier for each purchase order.
- **Product Name**: The name of the product being ordered.
- **Quantity**: The quantity of the product.
- **Price**: The price per unit of the product.
- **Total**: The total cost for the line item (Quantity × Price).
- **Customer Information**: Any additional details such as customer name, address, etc.

Ensure that the column names in the Excel file match those expected by the `main.py` script. 

## Template Document Explanation

The Word document template used by the program has placeholders for various fields that are populated with data from the Excel file. For example:

- `<<PO_Number>>`: Placeholder for the Purchase Order number.
- `<<Customer_Name>>`: Placeholder for the customer's name.
- `<<Product_Details>>`: A section where product details are added dynamically.

The template also includes a table where product rows are added for each item in the purchase order. The program ensures that the correct number of rows is added, with each row populated with data from the corresponding Excel file entry.

## Screenshots

### Template Document

![Template Document Screenshot](/main/TEMPLATE_Screenshot.png)


## Contributing

If you'd like to contribute to this project, please fork the repository and use a feature branch. Pull requests are warmly welcome.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
