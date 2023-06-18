# Web Page Semantic-analysis

#### This Python script performs text analysis on web pages by extracting content, calculating various metrics, and saving the results in an Excel file. It utilizes libraries such as numpy, requests, BeautifulSoup, pandas, re, and nltk to perform the necessary operations.

## Features
- Extracts web page content, including the title and main content.
- Cleans the extracted text by removing ```HTML``` tags, ```newlines```, and ```tabs```.
- Splits the content into sentences.
- Calculates various metrics, including ```character count```, ```sentence count```, ```polarity score```, ```subjectivity score```, ```readability indices```, etc.
- Uses predefined ```stop words``` to filter out irrelevant tokens.
- Determines ```positive and negative``` word counts.
- Writes the calculated metrics to an output ```Excel file```.

## Prerequisites

- Python 3 installed on your system.
- Required libraries: ```numpy```, ```requests```, ```beautifulsoup4```, ```pandas```, ```nltk```.

## Usage

- Prepare an Excel file ```(Input.xlsx)``` containing a column named URL with the URLs of the web pages you want to analyze. Optionally, you can include a column named ```URL_ID``` to identify each URL.
- Modify the script to provide the correct paths for input files and adjust any other necessary configurations.
- The script will retrieve the web page data, perform analysis, and save the results to an Excel file named ```Output Data Structure.xlsx```.
