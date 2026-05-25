Excel-to-Form Automation 🤖
Puppeteer-based automation that eliminated manual data entry from an Excel spreadsheet into a third-party management system, reducing processing time from ~30 seconds per record to under 1 second, across 100+ entries.
The Problem
A staff member had to manually transfer data from an Excel file into a web-based form- one person at a time. Each record required opening the form, filling multiple fields, and submitting. For 100 people, that's 50+ minutes of repetitive work.
The Solution
A Puppeteer script that reads each row from the Excel file and automatically fills and submits the form, handling the full flow without human intervention.
Results
BeforeAfterTime per record~30 seconds< 1 second100 records~50 minutes< 2 minutesHuman effortHighNone
Tech Stack

Python
Puppeteer (via Pyppeteer)
Excel / openpyxl — data source parsing

Run locally
bashgit clone https://github.com/Immerhaze/python-Automatizacion
cd python-Automatizacion
pip install -r requirements.txt
python main.py

Built by Nicode
