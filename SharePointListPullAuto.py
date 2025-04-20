import datetime
import pandas as pd
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

# Configuration
client_id = "your_client_id"
client_secret = "your_client_secret"
site_url = "https://yourtenant.sharepoint.com/sites/yoursite"
list_title = "YourListTitle"
excel_file = "output.xlsx"
last_checked_file = "last_checked.txt"

# Read last checked time
try:
    with open(last_checked_file, 'r') as f:
        last_checked_time = f.read().strip()
except FileNotFoundError:
    last_checked_time = "1900-01-01T00:00:00Z"

# Connect to SharePoint
ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

# Get the list
sp_list = ctx.web.lists.get_by_title(list_title)

# Get field names
fields = sp_list.fields.get().execute_query()
field_names = [field.internal_name for field in fields]

# Query new items
query = f"Created gt datetime'{last_checked_time}'"
l_items = sp_list.get_items().filter(query)
ctx.load(l_items)
ctx.execute_query()

# Collect data
data = []
for item in l_items:
    row = {field: item.properties.get(field, None) for field in field_names}
    data.append(row)

# Create DataFrame
new_df = pd.DataFrame(data)

# Read existing Excel file
try:
    existing_df = pd.read_excel(excel_file)
except FileNotFoundError:
    existing_df = pd.DataFrame()

# Append new data
combined_df = pd.concat([existing_df, new_df], ignore_index=True)

# Write to Excel
combined_df.to_excel(excel_file, index=False)

# Update last checked time
now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
with open(last_checked_file, 'w') as f:
    f.write(now)
