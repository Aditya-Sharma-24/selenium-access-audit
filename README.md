# Web Scraping Automation for User Access Audits
A Selenium-based Python tool to automate user group access audits across 70+ clients, saving over 100 hours/month.

## Overview
This script automates the extraction of user access data from an internal company portal. It offers three modes:
1. **User Group Scraping**: Extracts initial user group data into Excel.
2. **User Details (≤300 Groups)**: Scrapes detailed user info for smaller datasets.
3. **User Details (>300 Groups)**: Optimized for larger datasets with longer wait times.

- **Deployed**: August 2024
- **Impact**: Eliminated 100+ hours/month of manual audit work
- **Clients Supported**: 70+

## Features
- Manual login prompt via Tkinter for secure access.
- Scrapes paginated tables with Selenium (Edge WebDriver).
- Saves data to Excel with `openpyxl`.
- Menu-driven interface for selecting scraping mode.

## Tech Stack
- Python
- Selenium (Edge WebDriver)
- Openpyxl (Excel handling)
- Tkinter (GUI prompts)

## How to Run
1. **Prerequisites**:
   - Install dependencies: `pip install -r requirements.txt` (see below).
   - Microsoft Edge browser and Edge WebDriver installed.
2. **Setup**:
   - Update `URL` and `login_url` in `scraper.py` with your internal portal URL (currently placeholders).
   - Ensure `scraped_data.xlsx` exists for options 2 and 3 (output from option 1).
3. **Run**:
   - Execute: `python scraper.py`.
   - Choose an option:
     - `1`: Scrape user groups (manual login required).
     - `2`: Scrape user details (max 300 groups).
     - `3`: Scrape user details (>300 groups).
   - Follow prompts as needed.
4. **Output**:
   - Option 1: `scraped_data.xlsx` (user groups).
   - Options 2 & 3: `User_Details.xlsx` (detailed user info).

## Usage Notes
- **Option 1**: Scrapes user group data from paginated tables. Requires manual login.
- **Option 2**: Extracts detailed user info (e.g., name, email, role) for up to 300 groups. Faster execution.
- **Option 3**: Same as Option 2 but optimized for >300 groups with longer waits for stability.
- Internal tool—URL and XPaths are specific to the original portal.

## Status
Complete – Deployed August 2024
