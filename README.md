# 💰 ExpenseEase — Personal Expense Tracker

ExpenseEase is a personal expense tracking web application built using Python and Flask. It helps users record, manage, and analyze their daily expenses through a simple and user-friendly dashboard.

## ✨ Features

- **Dashboard:** View expense summaries and spending trends.
- **Expense Management:** Add, edit, and delete expense records.
- **Payment Methods:** Categorize transactions by Cash, UPI, Debit Card, Credit Card, Net Banking, and more.
- **Category Management:** Organize expenses into categories such as Food, Transport, Rent, Utilities, and Entertainment.
- **Search and Filters:** Find transactions and filter expenses by date or category.
- **Analytics:** Visualize spending patterns with charts.
- **Bulk Delete:** Select and delete multiple expense records.
- **CSV Export:** Export transaction records to CSV files.
- **Excel Export:** Download transaction records as Excel spreadsheets.
- **Profile and Settings:** Manage profile information and application preferences.
- **Local Database:** Store expense records using SQLite.

## 🛠️ Technologies Used

- **Backend:** Python, Flask
- **Database:** SQLite
- **ORM:** Flask-SQLAlchemy
- **Frontend:** HTML, Tailwind CSS, JavaScript
- **Charts:** Chart.js
- **Excel Export:** OpenPyXL
- **Development Environment:** Visual Studio Code

## 📁 Project Structure

```text
ExpenseEase/
├── app.py
├── templates/
│   ├── add_expense.html
│   ├── analytics.html
│   ├── base.html
│   ├── edit.html
│   ├── index.html
│   ├── profile.html
│   ├── settings.html
│   └── transactions.html
├── .gitignore
└── README.md
```

The SQLite database and Python virtual environment are stored locally and excluded from version control.

## ⚙️ Installation and Setup

### 1. Clone the repository

```bash
git clone https://github.com/Harshit1806/ExpenseEase.git
cd ExpenseEase
```

### 2. Create a virtual environment

**Windows:**

```bash
python -m venv venv
venv\Scripts\activate
```

If PowerShell blocks activation, use Command Prompt or run the virtual environment's Python executable directly.

### 3. Install dependencies

```bash
python -m pip install Flask Flask-SQLAlchemy openpyxl
```

### 4. Run the application

```bash
python app.py
```

Open the local URL displayed in the terminal. In the current project configuration, it is typically:

```text
http://127.0.0.1:4848
```

## 💾 Database

ExpenseEase uses SQLite to store expense records locally. The database is created or initialized by the application when it starts.

## 🔒 Security and Privacy

- Local database files are excluded from Git using `.gitignore`.
- Virtual environment files are excluded from version control.
- Never commit passwords, API keys, or other sensitive information.

## 🚀 Future Improvements

- Income tracking and balance summaries
- Monthly budget management
- Recurring transactions
- More detailed financial reports
- Improved mobile responsiveness

## 👨‍💻 Author

**Harshit**

GitHub: [@Harshit1806](https://github.com/Harshit1806)

## 📄 License

No license has been added to this repository yet.
