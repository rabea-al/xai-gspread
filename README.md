
<p align="center">
  <a href="https://github.com/XpressAI/xircuits/tree/master/xai_components#xircuits-component-library-list">Component Libraries</a> •
  <a href="https://github.com/XpressAI/xircuits/tree/master/project-templates#xircuits-project-templates-list">Project Templates</a>
  <br>
  <a href="https://xircuits.io/">Docs</a> •
  <a href="https://xircuits.io/docs/Installation">Install</a> •
  <a href="https://xircuits.io/docs/category/tutorials">Tutorials</a> •
  <a href="https://xircuits.io/docs/category/developer-guide">Developer Guides</a> •
  <a href="https://github.com/XpressAI/xircuits/blob/master/CONTRIBUTING.md">Contribute</a> •
  <a href="https://www.xpress.ai/blog/">Blog</a> •
  <a href="https://discord.com/invite/vgEg2ZtxCw">Discord</a>
</p>

<p align="center"><i>Xircuits Component Library for GSpread – Simplify your Google Sheets operations directly in Xircuits workflows.</i></p>

---

## Xircuits Component Library for GSpread

Integrate Google Sheets functionalities into Xircuits workflows effortlessly. Perform operations like spreadsheet creation, cell updates, and worksheet management using the GSpread library.


## Table of Contents

- [Preview](#preview)
- [Prerequisites](#prerequisites)
- [Main Xircuits Components](#main-xircuits-components)
- [Try the Examples](#try-the-examples)
- [Installation](#installation)


## Preview

### SimpleGSpread Example 

<img src="https://github.com/user-attachments/assets/fee8c1ad-68af-4243-bdc0-5d5e9c0d1046" alt="SimpleGSpread_example" />

### SimpleGSpread Result

<img src="https://github.com/user-attachments/assets/52fd64bb-7949-468f-ac98-b5d14c43029b" alt="SimpleGSpread" />

## Prerequisites

Before you begin, you will need:

1. Python 3.9+.
2. Xircuits installed.
3. Google Service Account JSON key for authentication.


## Main Xircuits Components

### GspreadAuth Component:
Authenticates using a Google Service Account JSON key and initializes a GSpread client.

<img src="https://github.com/user-attachments/assets/049d0621-fda9-478d-b7f1-0e47f1664198" alt="GspreadAuth" width="200" height="75" />

### OpenSpreadsheet Component:
Opens a Google Spreadsheet by its title and selects a specified worksheet or defaults to the first worksheet.

<img src="https://github.com/user-attachments/assets/33d78200-17d4-4416-8ca7-7089d699a165" alt="OpenSpreadsheet" width="200" height="150" />

### OpenSpreadsheetFromUrl Component:
Opens a Google Spreadsheet using its URL and selects a worksheet.

### OpenWorksheet Component:
Selects a worksheet by its title or defaults to the first worksheet.

### CreateSpreadsheet Component:
Creates a new Google Spreadsheet and optionally shares it with an email.

### CreateWorksheet Component:
Adds a new worksheet to an existing spreadsheet with customizable rows and columns.

### UpdateRange Component:
Updates a range of cells in the worksheet with specified values.

### GetAllValues Component:
Retrieves all values from a worksheet as a list of lists.

### GetAllRecords Component:
Retrieves all records from a worksheet as a list of dictionaries.

### FindAllStringMatches Component:
Finds all cells containing a specific string value.

### FindAllRegexMatches Component:
Finds all cells matching a specified regular expression.


## Try the Examples

We have provided example workflows to help you get started with the GSpread component library. These examples demonstrate how to authenticate, read, and update Google Sheets within a Xircuits workflow.

### SimpleGSpread
This example demonstrates how to authenticate, read a cell's value, and update it with new data.


## Installation
To use this component library, ensure that you have an existing [Xircuits setup](https://xircuits.io/docs/main/Installation). You can then install the GSpread library using the [component library interface](https://xircuits.io/docs/component-library/installation#installation-using-the-xircuits-library-interface), or through the CLI using:

```bash
xircuits install gspread
```

Alternatively, you can install it manually by cloning and installing it:

```bash
# base Xircuits directory
git clone https://github.com/XpressAI/xai-gspread xai_components/xai_gspread
pip install -r xai_components/xai_gspread/requirements.txt
```

### Authentication and Service Credentials Setup

To access spreadsheets via Google Sheets API, you need to authenticate and authorize your application. If you plan to access spreadsheets on behalf of a bot account, use a Service Account. Below are the steps to set up service credentials.

1. Enable API Access:
    1. Go to Google Developers Console and create a new project (or select an existing one).
    2. In "Search for APIs and Services," enable both "Google Drive API" and "Google Sheets API."
2. Create Service Account Credentials:
    1. Navigate to "APIs & Services > Credentials" and click "Create credentials > Service account key."
    2. Fill out the form and select JSON key type.
    3. Download the JSON file containing your credentials.
3. Share Spreadsheet Access:
    - Share your spreadsheet with the `client_email` listed in the JSON file.
4. Configure Credentials:
    - Place the downloaded JSON file in a location accessible to your application.
    - Optionally, move the file to `~/.config/gspread/service_account.json` (Linux/Mac) or `%APPDATA%\gspread\service_account.json` (Windows).

For more details, refer to the [GSpread Authentication Guide](https://docs.gspread.org/en/latest/oauth2.html).

## Notes
1. Ensure the Google Service Account is shared with the spreadsheets you intend to access.
2. Keep your JSON key secure and do not expose it publicly.
