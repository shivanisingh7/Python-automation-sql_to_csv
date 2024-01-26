import win32com.client
import pandas as pd
from datetime import datetime
#query to get the sever name is SELECT @@SERVERNAME is SSMS
#have used double backlashes in order to avoid escape characters
#have used Microsoft OLE DB Driver for SQL Server
#for database have used AdventureWorks databases are sample databases that were originally published by Microsoft to show how to design a SQL Server database using SQL Server 2008

Connection_string = (
    'Provider=MSOLEDBSQL;'
    'Data Source=SHIVANI\\SQLEXPRESS;'
    'Initial Catalog=AdventureWorks2022;'
    'Integrated Security=SSPI;'
)

conn = win32com.client.Dispatch('ADODB.Connection')
conn.Open(Connection_string)

# Use the connection to execute a SQL query
sql_query = "SELECT BusinessEntityID, PasswordHash, rowguid FROM Person.Password"
recordset = win32com.client.Dispatch('ADODB.Recordset')
recordset.Open(sql_query, conn)

columns = [field.Name for field in recordset.Fields]
print(columns)  # Print the field names to check

# Load the data into a list
data = []
while not recordset.EOF:
    row_data = [recordset.Fields.Item(field).Value for field in columns]
    data.append(row_data)
    recordset.MoveNext()

# Close the connections
recordset.Close()
conn.Close()

# Create the DataFrame
DF = pd.DataFrame(data, columns=columns)

# Save DataFrame to CSV
DF.to_csv(datetime.now().strftime("%Y-%m-%d_%I-%M-%S-%p") + 'Sql_user_password_Data.csv', index=False)