import pyodbc
from tkinter import filedialog
import tkinter as tk

def select_database_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select Access Database File",
        filetypes=[("Access Database", "*.mdb;*.accdb"), ("All Files", "*.*")],
    )

    return file_path

# Get the path to the Access database file from the user
database_file_path = select_database_file()

# Establish a connection to the Access database
conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={database_file_path};"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Get the column names in the 'Check_Ns_Close' table
cursor.execute("SELECT * FROM Check_Ns_Close")
columns = [column[0] for column in cursor.description]
columns_to_analyze = columns[1:]  # Exclude the first field

# Create a new table to store the summarized information
create_table_query = "CREATE TABLE Check_Ns_Process (FieldName TEXT, Total INTEGER, "
for i in range(-1, 31):
    create_table_query += f"{i} INTEGER, "
create_table_query = create_table_query.rstrip(', ') + ")"
cursor.execute(create_table_query)

# Count occurrences of each value in each column and insert into the Check_Ns_Process
for col in columns_to_analyze:
    count_query = (
        f"INSERT INTO Check_Ns_Process (FieldName, Total, {', '.join(map(str, range(-1, 31)))}) "
        f"SELECT '{col}', SUM(IIF([{col}]>0, 1, 0)), {', '.join([f'SUM(IIF([{col}]={i}, 1, 0))' for i in range(-1, 31)])} FROM Check_Ns_Close"
    )
    cursor.execute(count_query)

# Commit changes and close the connection
conn.commit()
conn.close()
