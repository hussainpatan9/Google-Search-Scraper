# Google Search Scraper

This Python script allows you to perform Google searches using the Custom Search JSON API, retrieve the results, and save them to an Excel file. The script provides a graphical user interface (GUI) built using Tkinter for easy interaction.

## Prerequisites

Before using the script, make sure you have the following:

- Python installed (version 3.6 or higher)
- The required Python libraries installed (`openpyxl`, `tkinter`, `requests`)

You can install the required libraries using:

```bash
pip install openpyxl tkinter requests
```

## Getting Started

1. Clone the repository to your local machine:

```bash
git clone https://github.com/hussainpatan9/Google-Search-Scraper.git
```

2. Navigate to the project directory:

```bash
cd google-search-scraper
```

3. Open the `config.json` file and add your Google Custom Search JSON API key and Search Engine ID. If you don't have them, you can obtain them from the [Google Custom Search JSON API](https://developers.google.com/custom-search/v1/overview).

```json
{
  "API_KEY": "YOUR API KEY HERE",
  "SEARCH_ENGINE_ID": "YOUR SEARCH ENGINE ID HERE"
}
```

4. Run the script:

```bash
python google_search_scraper.py
```

## Usage

1. The GUI will prompt you to select a keywords file. This file should contain one keyword per line.
2. Choose the output folder where the Excel file with the search results will be saved.
3. Click the "Run" button to start the Google searches.

The script will create an Excel file with the search results, including the keyword, rank, title, URL, and description.

## Important Note

- Make sure to keep your API key and Search Engine ID confidential. Do not expose them on public repositories or share them with unauthorized users.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the creators of the Python libraries used in this script.
