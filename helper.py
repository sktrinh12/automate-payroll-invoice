from datetime import datetime
import os.path
import glob
import argparse

def pprint(message, separator="="):
    width = len(message)
    print(f"{message}\n{separator * width}\n")


def count_invoices(base_directory):
    pattern = os.path.join(os.getenv('BASE_DIR'), "**", f"{os.getenv('PREFIX')}_*.xlsx")
    invoice_files = glob.glob(pattern, recursive=True)
    invoice_count = len(invoice_files) + 1
    pprint(f"invoice count: {invoice_count}")
    return invoice_count


def parse_arguments():
    parser = argparse.ArgumentParser(description='Process a date string in the format DD-MM-YYYY or MM-DD-YYYY.')
    # required date argument
    parser.add_argument('-d', '--date', type=str, required=True, help='Required date string (e.g., 15-10-2024 or 10-15-2024)')
    # optional argument to skip drafting of email
    parser.add_argument('-e', '--email', action='store_false', help='Optional flag to trigger drafting an email')
    # optional argument to skip uploading to sqlite3 db
    parser.add_argument('-a', '--add_sqlite3', action='store_false', help='Optional flag to trigger upload to sqlite3 db')
    args = parser.parse_args()

    try:
        parsed_date = datetime.strptime(args.date, '%d-%m-%Y')
    except ValueError:
        try:
            parsed_date = datetime.strptime(args.date, '%m-%d-%Y')
        except ValueError:
            parser.error("Date must be in the format DD-MM-YYYY or MM-DD-YYYY.")

    return parsed_date.day, parsed_date.month, parsed_date.year, args.email, args.add_sqlite3


def load_env_file(env_file_path):
    """Load key-value pairs from a .env file into environment variables."""
    try:
        with open(env_file_path, 'r') as file:
            for line in file:
                # Skip empty lines and comments
                line = line.strip()
                if line and not line.startswith('#'):
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip()
        pprint("Environment variables loaded successfully.")
    except FileNotFoundError:
        pprint(f">>>.env file not found: {env_file_path}")
    except Exception as e:
        pprint(f">>>An error occurred while loading the .env file: {e}")
