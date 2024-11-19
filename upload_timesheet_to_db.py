import pandas as pd
import sqlite3
import argparse
import os.path

def extract_data_from_xlsx(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension == '.xlsx':
        df = pd.read_excel(file_path)
    elif file_extension == '.csv':
        df = pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file type. Only .xlsx and .csv are supported.")
    
    columns = ["ID", "Date/time", "Decimal hours"]
    
    if not all(col in df.columns for col in columns):
        raise ValueError(f"File must contain {', '.join(columns)} columns.")
    
    df["Date"] = pd.to_datetime(df["Date/time"], format="%m/%d/%Y %I:%M%p", errors='coerce')
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")
    df["File_Path"] = file_path

    df = df[["ID", "Date", "Decimal hours", "File_Path"]]
    df.columns = ["ID", "DATE_TIME", "DECIMAL_HOURS", "FILE_PATH"]

    return df

def upload_to_sqlite(df, db_name="timesheet.db"):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS RCH_TIMESHEET (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            DATE_TIME TEXT,
            DECIMAL_HOURS REAL,
            FILE_PATH TEXT
        )
    ''')
    
    for _, row in df.iterrows():

        try:
            cursor.execute('''
                INSERT INTO RCH_TIMESHEET (ID, DATE_TIME, DECIMAL_HOURS, FILE_PATH)
                VALUES (?, ?, ?, ?)
            ''', (row["ID"], row["DATE_TIME"], row["DECIMAL_HOURS"], row["FILE_PATH"]))
        except sqlite3.IntegrityError as e:
            print(f"Primary key error on row {_}: {row.to_dict()}")
            print(f"Error message: {e}")
    
    conn.commit()
    conn.close()
    print(f"Data from ({df['FILE_PATH'].iloc[0]}) uploaded successfully.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Upload data from local Excel timesheet file to local SQLITE3 database"
    )
    parser.add_argument(
        "file_path", type=str, help="Path to the local Excel timesheet file"
    )
    args = parser.parse_args()

    if not os.path.isfile(args.file_path):
        raise FileNotFoundError(
            f"The file '{args.file_path}' does not exist or is not a valid file path."
        )
    df = extract_data_from_xlsx(args.file_path)
    upload_to_sqlite(df)
