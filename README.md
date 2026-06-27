# registration-extractor-tool

A tool to extract, clean and sync event registrations from CSV exports to Google Sheets.

---

## What it does

1. Reads a CSV file exported from your registration system
2. Removes duplicate registrations (keeps the most recent one per person)
3. Filters out blocked Telegram accounts (blacklist)
4. Sorts registrations by payment status
5. Saves the result as an XLSX file
6. Uploads the data to a Google Sheet with checkboxes

---

## Requirements

- Python 3.10 or higher
- A Google account with access to your Google Sheet
- A Google Cloud project with the Google Sheets API enabled

---

## Setup

### 1. Clone the repository

```bash
git clone <repo-url>
cd registration-extractor-tool
```

### 2. Create a virtual environment and install dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Set up Google Sheets API

**Step 1. Create a Google Cloud project**

- Go to [console.cloud.google.com](https://console.cloud.google.com)
- Click the project dropdown at the top → **New Project**
- Give it any name, click **Create**
- Make sure the new project is selected in the dropdown

**Step 2. Enable the Google Sheets API**

- In the left menu go to **APIs & Services → Library**
- Search for **Google Sheets API**
- Click on it and press **Enable**

**Step 3. Configure the OAuth consent screen**

- Go to **APIs & Services → OAuth consent screen**
- Choose **External** and click **Create**
- Fill in the **App name** (anything) and your **Gmail** as the support email
- Scroll to the bottom and click **Save and Continue** through all steps
- On the last step (**Test users**) click **+ Add users** and add your Gmail address
- Click **Save and Continue**

**Step 4. Create OAuth credentials**

- Go to **APIs & Services → Credentials**
- Click **+ Create Credentials → OAuth 2.0 Client ID**
- Choose **Desktop app** as the application type
- Click **Create**
- Click the **Download JSON** button (the download icon on the right)
- Rename the downloaded file to `credentials.json`
- Move it to the project folder (next to `main.py`)

### 4. Configure the project

Copy the example env file and fill in your values:

```bash
cp .env.example .env
```

Open `.env` and set:

```
SPREADSHEET_ID=your_spreadsheet_id_here
SHEET_NAME=your_sheet_name_here
```

You can find the Spreadsheet ID in the URL of your Google Sheet:  
`https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit`

---

## Usage

```bash
source venv/bin/activate
python3 main.py input.csv output.xlsx
```

On the first run, a browser window will open asking you to log in to Google and grant access.  
After that, the token is saved to `token.pickle` and you won't need to log in again.

### Skip Google Sheets upload

If you only want the XLSX file without uploading to Google Sheets:

```bash
python3 main.py input.csv output.xlsx --no-sheets
```

---

## Filters

All filtering, status mapping, sort order and colors are configured in `filters.yaml` — no need to touch the code.

```yaml
# Colors for the "Статус регистрации" cell in the XLSX file (hex, no #)
xlsx_colors:
  Оплатил: 90EE90
  Зарегистрировался: ADD8E6
  Надо напомнить: FF6B6B

blacklist:
  # Exact Telegram handles (without @, case-insensitive)
  exact:
    - username

  # Block all handles that start with these strings
  prefixes:
    - spamuser

  # Regular expressions matched against the full handle
  patterns:
    - "test\\d+"

# Registrations with these CSV statuses are excluded from output entirely
status_filter:
  exclude:
    - Отменено

# Maps original CSV status to the output label
status_map:
  Оплачено: Оплатил
  Ожидает оплаты: Зарегистрировался
  Просрочено: Надо напомнить

# Output sort order (lower number = higher in the list)
sort_order:
  Оплачено: 0
  Ожидает оплаты: 1
  Просрочено: 2
```

You can also use a custom filters file:

```bash
python3 main.py input.csv output.xlsx --filters my_filters.yaml
```

---

## Output columns

| Column | Description |
|---|---|
| Статус регистрации | Mapped status: Оплатил / Зарегистрировался / Напомнили / Отдали промокод |
| Сайт актуален? | Checkbox (checked by default) |
| ID регистрации | Registration UUID |
| ФИО | Full name |
| Email | Email address |
| Telegram | Telegram handle |
| Волна | Registration wave |
| Билет | Ticket type |
| Статус | Original status from CSV |
| Тип оплаты | Payment method |
| Полная сумма | Full price |
| Скидка | Discount |
| Итого | Total amount due |
| Время регистрации | Registration timestamp |
| Время подтверждения | Payment confirmation timestamp |
| Получатель (имя) | Transfer recipient name |
| Получатель (банк) | Recipient bank |
| Реквизиты перевода | Card or phone number |

### Status mapping

| CSV status | Output status |
|---|---|
| Оплачено | Оплатил |
| Ожидает оплаты | Зарегистрировался |
| Просрочено | Надо напомнить |
| Отменено | filtered out, not included in output |

### Sort order

Оплатил → Зарегистрировался → Надо напомнить

---

## Project structure

```
.
├── main.py            # main script
├── filters.yaml       # all filters, statuses, blacklist and colors
├── requirements.txt   # dependencies
├── .env.example       # environment variables template
├── .env               # your local config (never commit this)
├── credentials.json   # Google OAuth credentials (never commit this)
├── token.pickle       # auto-generated after first login (never commit this)
└── .gitignore
```

---

## Security

Never commit `.env`, `credentials.json` or `token.pickle` — they contain your personal access tokens and API secrets. They are already listed in `.gitignore`.
