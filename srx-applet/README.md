# SRx SoS Tracks Applet

A Google Apps Script web app that reads HCC V2MOM data and renders an SoS track dashboard.

## Setup

1. Go to [script.google.com](https://script.google.com) → **New Project**
2. Replace the default `Code.gs` with the contents of `Code.gs` from this folder
3. Click **+** next to Files → **HTML** → name it `Index` → paste contents of `Index.html`
4. Click **Deploy** → **New Deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone with Salesforce account** (or "Anyone" for wider access)
5. Click **Deploy** and authorize when prompted
6. Copy the deployment URL — that's your applet link

## Data Source

- **Spreadsheet:** [HCC FY27 V2MOM](https://docs.google.com/spreadsheets/d/1ufYd6q4YhSUDFFRJpXgjuwoPcKKugd0SB0oOHlj06fM/edit?gid=649674725#gid=649674725)
- **Rows:** 25–33
- **Columns:**
  - C = Measure name
  - E = Owner
  - L = Status
  - M = Progress/comments
  - N = Notes/updates

## Updating

To refresh data, just reload the web app URL — it reads live from the spreadsheet each time.

To change which rows/columns are read, edit `Code.gs` and redeploy.
