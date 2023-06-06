# Xircuits Google Spreadsheet Component Library


Component library based on the [GSpread](https://github.com/burnash/gspread) library.

![image](https://github.com/XpressAI/xai-gspread/assets/68586800/b4c61851-47bc-4e89-b5de-5d9c4aa985b1)


## Important Concepts
- There are gspread clients `gc`, spreadsheets `sh`, as well as `worksheet`s.
- A single gspread client can access multiple spreadsheets provided it has the permission to, and spreadsheet can contain multiple worksheets.
- The newest instance of `gc`, `sh` or `worksheet` will be set on the `ctx` by default so you do not need to link each port.
- However you can also refer to a previous `gc`, `sh`, or `worksheet` by linking their ports.


## Authentication

[[Source](https://docs.gspread.org/en/latest/oauth2.html)]

To access spreadsheets via Google Sheets API you need to authenticate and authorize your application. If you plan to access spreadsheets on behalf of a bot account use Service Account. If you’d like to access spreadsheets on behalf of end users (including yourself) use OAuth Client ID. For this demo, we're using the Service Credentials.

### Getting Service Credentials

A service account is a special type of Google account intended to represent a non-human user that needs to authenticate and be authorized to access data in Google APIs.

Since it’s a separate account, by default it does not have access to any spreadsheet until you share it with this account. Just like any other Google account.


1. Enable API Access.
    1. Head to Google Developers Console and create a new project (or select the one you already have).
    2. In the box labeled “Search for APIs and Services”, search for “Google Drive API” and enable it.
    3. In the box labeled “Search for APIs and Services”, search for “Google Sheets API” and enable it.
2. Go to “APIs & Services > Credentials” and choose “Create credentials > Service account key”.
3. Fill out the form
4. Click “Create” and “Done”.
5. Press “Manage service accounts” above Service Accounts.
6. Press on ⋮ near recently created service account and select “Manage keys” and then click on “ADD KEY > Create new key”.
7. Select JSON key type and press “Create”.
You will automatically download a JSON file with credentials. It may look like this:
    ```
    {
        "type": "service_account",
        "project_id": "api-project-XXX",
        "private_key_id": "2cd … ba4",
        "private_key": "-----BEGIN PRIVATE KEY-----\nNrDyLw … jINQh/9\n-----END PRIVATE KEY-----\n",
        "client_email": "473000000000-yoursisdifferent@developer.gserviceaccount.com",
        "client_id": "473 … hd.apps.googleusercontent.com",
        ...
    }
    ```
  Remember the path to the downloaded credentials file. Also, in the next step you’ll need the value of client_email from this file.

8. **Very important!** Go to your spreadsheet and share it with a client_email from the step above. Just like you do with any other Google account. If you don’t do this, you’ll get a gspread.exceptions.SpreadsheetNotFound exception when trying to access this spreadsheet from your application or a script.
9. Place the credentials in a place that jupyterlab can access. You can then pass the path in the `GSpreadAuth` component. For more permanent solutions, you can move the downloaded file to ~/.config/gspread/service_account.json. Windows users should put this file to %APPDATA%\gspread\service_account.json.
