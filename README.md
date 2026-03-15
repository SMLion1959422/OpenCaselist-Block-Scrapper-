**# OpenCaselist-Block-Scrapper**
**Background:**
Built a Python automation tool that collects open-source debate round files from OpenCaselist, filters them by team, school, recency, or topic, and compiles them into a single research packet while preserving document formatting.
I coded this tool because I saw my debate event(public forum) was heavily reliant on evidence and prep. Teams that payed money for coaching and prep were a lot prepared in tournaments and was a huge reason I have lost rounds. That is why I chose to build this tool and push the debate space as sinething equitable and open for all. 

## Features

- Scrape round files from OpenCaselist by:
  - specific teams
  - entire schools
  - recently uploaded rounds
  - topic keywords
- Filter rounds using custom keyword matching
- Cache API responses and downloaded files to reduce duplicate requests
- Merge source documents into a single compiled research packet
- Preserve original DOCX formatting during document merge
- Export output as DOCX and optionally convert to PDF

## Tech Stack

- Python
- Requests
- python-docx
- docx2pdf
- OpenCaselist API

## How It Works

1. Authenticate using your OpenCaselist token
2. Select a target mode:
   - `teams`
   - `school`
   - `recent`
   - `topic`
3. Fetch matching rounds from the API
4. Remove duplicate files
5. Download and cache source documents
6. Merge them into one formatted output packet
7. Save the result as DOCX or PDF

**Installation**
pip install requests python-docx docx2pdf

**Usage**
Open the script and replace:

CASELIST_TOKEN = "REPLACE_ME_WITH_YOUR_TOKEN"

Run the program:

python Caselistscrapper.py

