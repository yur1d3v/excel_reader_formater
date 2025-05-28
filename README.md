# Excel Reader & Formatter

This project reads an Excel file, processes its data (including calculating a new column), saves the modified file, and automatically sends it by email.

## Features

- Reads all sheets from `data.xlsx`
- Strips whitespace from column names
- Converts columns to numeric and calculates a new column `Valor Total`
- Saves the result as `data_modified.xlsx`
- Sends the new file as an email attachment automatically

## Requirements

- Python 3.8+
- See `requirements.txt` for dependencies

## Setup

1. **Clone the repository:**
   ```sh
   git clone https://github.com/yourusername/excel_reader_formater.git
   cd excel_reader_formater
   ```

2. **Install dependencies:**
   ```sh
   pip install -r requirements.txt
   ```

3. **Create a `.env` file** in the project folder with your email credentials:
   ```
   EMAIL_ADDRESS=your_email@gmail.com
   EMAIL_PASSWORD=your_app_password
   TO_ADDRESS=recipient@example.com
   ```

4. **Place your `data.xlsx` file in the project folder.**

5. **Run the script:**
   ```sh
   python excel_reader.py
   ```

## Notes

- For Gmail, use an [App Password](https://support.google.com/accounts/answer/185833?hl=en) if you have 2FA enabled.
- The script will raise an error if `data_modified.xlsx` already exists.

---

**Replace `yourusername` in the clone URL with your actual GitHub username.**
