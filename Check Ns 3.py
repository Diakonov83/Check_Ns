import pyodbc

# Establish a connection to the Access database
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=D:\LocalData\VMShare\0_Work_Temp\23132-St. Charles Health System\MASTER.mdb;"
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Get the column names in the 'Counting' table
cursor.execute("SELECT * FROM Counting")
columns = [column[0] for column in cursor.description]
columns_to_analyze = columns[1:]  # Exclude the first field

# Create a new table to store the summarized information
create_table_query = "CREATE TABLE SummaryTable (FieldName TEXT, Total INTEGER, "
for i in range(-1, 31):
    create_table_query += f"{i} INTEGER, "
create_table_query = create_table_query.rstrip(', ') + ")"
cursor.execute(create_table_query)

# Count occurrences of each value in each column and insert into the SummaryTable
for col in columns_to_analyze:
    count_query = (
        f"INSERT INTO SummaryTable (FieldName, Total, {', '.join(map(str, range(-1, 31)))}) "
        f"SELECT '{col}', SUM(IIF([{col}]>0, 1, 0)), {', '.join([f'SUM(IIF([{col}]={i}, 1, 0))' for i in range(-1, 31)])} FROM Counting"
    )
    cursor.execute(count_query)

# Commit changes and close the connection
conn.commit()
conn.close()
