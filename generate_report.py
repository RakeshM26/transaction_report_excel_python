import psycopg2
import openpyxl
from datetime import datetime

# Database connection setup
def connect_db():
    try:
        connection = psycopg2.connect(
            host="localhost",   # Your DB host
            database="massimopro_dev", # Your DB name
            user="rakesh",   # Your DB username
            password="password" # Your DB password
        )
        return connection
    except Exception as e:
        print(f"Error connecting to the database: {e}")
        return None

# Function to get transactions by firm on a particular day
def get_transaction_data(connection, date):
    try:
        with connection.cursor() as cursor:
            query = """
            SELECT id, 
                   (SELECT COUNT(*) FROM invoice WHERE DATE(invoice.record_date) = %s AND invoice.firm_id = f.id) AS invoice_count,
                   (SELECT SUM(net_total) FROM invoice WHERE DATE(invoice.record_date) = %s AND invoice.firm_id = f.id) AS invoice_total,
                   (SELECT COUNT(*) FROM grn WHERE DATE(grn.record_date) = %s AND grn.firm_id = f.id) AS grn_count,
                   (SELECT COUNT(*) FROM purchase_order WHERE DATE(purchase_order.record_date) = %s AND purchase_order.firm_id = f.id) AS po_count
            FROM company f;
            """
            cursor.execute(query, (date, date, date, date))
            return cursor.fetchall()
    except Exception as e:
        print(f"Error fetching data: {e}")
        return []

# Function to create Excel file with transaction data
def create_excel_report(data, date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Report {date}"

    # Headers
    ws.append(["Firm ID", "Invoice Count", "Invoice Total", "GRN Count", "PO Count"])

    # Data rows
    for row in data:
        ws.append(row)

    # Save the Excel file
    filename = f"transaction_report_{date}.xlsx"
    wb.save(filename)
    print(f"Report saved as {filename}")

# Main function
def main():
    date = input("Enter the date (YYYY-MM-DD) for which you want the report: ")
    
    # Connect to the database
    connection = connect_db()
    if connection is None:
        return

    # Get transaction data
    transaction_data = get_transaction_data(connection, date)

    # Create Excel report
    create_excel_report(transaction_data, date)

    # Close the database connection
    connection.close()

if __name__ == "__main__":
    main()
