import snowflake.connector
import pandas as pd

# Snowflake connection settings
SNOWFLAKE_ACCOUNT = '<your_account_name>.snowflakecomputing.com'
SNOWFLAKE_USER = '<your_username>'
SNOWFLAKE_PASSWORD = '<your_password>'
SNOWFLAKE_WAREHOUSE = '<your_warehouse>'
SNOWFLAKE_DATABASE = '<your_database>'
SNOWFLAKE_SCHEMA = '<your_schema>'

# Establish connection to Snowflake
conn = snowflake.connector.connect(
    user=SNOWFLAKE_USER,
    password=SNOWFLAKE_PASSWORD,
    account=SNOWFLAKE_ACCOUNT,
    warehouse=SNOWFLAKE_WAREHOUSE,
    database=SNOWFLAKE_DATABASE,
    schema=SNOWFLAKE_SCHEMA
)

# Fetch data from Snowflake
query = "SELECT * FROM <your_table>"
cur = conn.cursor()
cur.execute(query)
rows = cur.fetchall()
column_names = [col[0] for col in cur.description]

# Load data into a pandas DataFrame
data = pd.DataFrame(rows, columns=column_names)

# Write data to Excel
excel_file = 'snowflake_data.xlsx'
data.to_excel(excel_file, index=False, engine='openpyxl')

# Close Snowflake connection
cur.close()
conn.close()

print(f"Data successfully written to {excel_file}")
