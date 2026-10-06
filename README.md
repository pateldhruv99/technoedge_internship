# TechnoEdge Blog Scraper

Python automation for extracting blog content from the TechnoEdge Blog Insights website and exporting the results to Excel.

## Features

- Automatically detects blog archive pages
- Extracts blog title, date, excerpt, thumbnail and URL
- Opens individual blog articles
- Extracts author, publication date, headings, tables and full article content
- Removes duplicate article URLs
- Saves progress using a checkpoint file
- Supports restarting after interruptions
- Exports structured data to Excel

## Website

https://www.technoedgels.com/blog-insights/

## Technologies

- Python
- Requests
- BeautifulSoup
- Pandas
- OpenPyXL
- lxml

## Installation

```bash
python -m venv venv
pip install -r requirements.txt