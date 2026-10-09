# 💸 Personal Expense Tracker

A simple, Python-based expense tracker that helps you log and visualize your daily spending. The project stores your data locally in `.csv` and `.xlsx` formats.

This project includes two ways to run the tracker:
1. **Interactive Command-Line Interface (CLI)**
2. **Modern Web Application (Streamlit)**

---

## 🚀 Features

- **Add New Expenses:** Easily log your expenses with a name, amount, category, and date.
- **View Summary:** See how much you've spent, view your remaining budget, and calculate your daily budget limits.
- **Data Visualizations:** Beautiful pie charts and bar charts to break down your expenses by category.
- **Edit & Delete:** Fix mistakes by editing or deleting previous entries.
- **Auto-Archiving:** Automatically moves past month's data into separate Excel sheets (`expense_history.xlsx`) when a new month begins.

---

## 🛠️ Setup & Installation

Make sure you have Python installed. Then, install the required libraries:

```bash
pip install pandas openpyxl matplotlib streamlit
```

---

## 💻 How to Run

### Option 1: Web Application (Recommended)
Enjoy a modern UI right in your browser (and accessible on your phone on the same Wi-Fi network).

```bash
cd "Expense tracker"
python -m streamlit run app.py
```

### Option 2: Command-Line Interface (CLI)
A fast, lightweight terminal app.

```bash
cd "Expense tracker"
python expense_tracker.py
```

---

## 📂 File Structure

- `Expense tracker/app.py` - The Streamlit web application.
- `Expense tracker/expense_tracker.py` - The CLI application.
- `Expense tracker/expense.py` - The Expense data model class.
- `expenses.csv` - Current month's active expenses (Ignored by git for privacy).
- `expense_history.xlsx` - Historical archives of past months (Ignored by git for privacy).
