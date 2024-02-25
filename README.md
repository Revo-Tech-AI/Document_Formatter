# HTML to Word Document Converter

This project provides a Python script that automates the conversion of HTML content, particularly formatted ChatGPT outputs, into Microsoft Word documents. This utility is designed to save time and effort that would otherwise be spent on manually formatting ChatGPT outputs in Word. The script parses HTML, retaining styles such as font size, family, weight, and color, and then applies these styles to the content in a newly created Word document.

## Requirements

- Python 3.x
- BeautifulSoup4
- python-docx

## Installation

Before you can run the script, you need to ensure you have Python installed on your machine and then install the required Python packages.

1. Install BeautifulSoup4: pip install beautifulsoup4 

2. Install python-docx: pip install python-docx 

## Usage

To use the script, follow these steps:

1. Prepare your HTML content that you wish to convert. This can be the output from ChatGPT or any other HTML content.

2. Save the script to a file, for example, `html_to_word.py`.

3. Run the script, passing the HTML content and the desired output Word document file name as arguments. Here's an example command:

python html_to_word.py your_input.html your_output.docx 

### Example

Given the HTML content in a file named `example.html`, you can convert it to a Word document named `Workshop_Concepts.docx` by running:

python html_to_word.py example.html Workshop_Concepts.docx


This script simplifies the process of converting HTML formatted text into a Word document, preserving the styling and structure, making it an invaluable tool for anyone looking to automate their documentation workflow.


