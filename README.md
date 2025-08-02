# Automate Excel and Google Sheets Tasks with Python

This project demonstrates how to automate tasks with Microsoft Excel and Google Sheets using Python. It includes practical examples for reading, writing, formatting, and syncing spreadsheet data, complete with error handling and logging.

Read the full tutorial [here]().

## 🚀 Features

- Read and write Excel files using `pandas` and `openpyxl`
- Style Excel headers
- Authenticate with Google Sheets using a service account
- Upload and update Google Sheets with `gspread` and `gspread_dataframe`
- Error handling and logging with `logging`
- Modular Python code for easy reuse

## 📦 Requirements

- Python 3.7+
- `pandas`
- `openpyxl`
- `gspread`
- `gspread_dataframe`
- `oauth2client`

Install dependencies:

```bash
pip install pandas openpyxl gspread gspread_dataframe oauth2client
```

## 🔐 Google Sheets API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project and enable the **Google Sheets API**
3. Create a **Service Account**, generate a key in **JSON** format, and download it as `credentials.json`
4. Share your target Google Sheet with the service account email

## 📂 Project Structure

```
excel_google_automation/
├── main.py              # Main automation script
├── credentials.json     # Google API credentials (not included)
├── automation.log       # Generated log file
├── students.xlsx        # Example Excel file (generated)
└── README.md            # Project documentation
```

## 🧪 How to Use

1. Place your `credentials.json` in the project directory
2. Run the script:

```bash
python main.py
```

3. Check the generated Excel file and your Google Sheet titled **"Students Report"**

## 🛠 Customize

You can modify the `main()` function in `main.py` to:

- Load data from a CSV or database
- Update different Google Sheets
- Schedule the script using `cron` or Windows Task Scheduler

## 🧾 License

MIT

---

> Tutorial by [Djamware.com]() - _Automate Excel and Google Sheets Tasks with Python: Practical Examples_
