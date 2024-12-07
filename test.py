import mstarpy
import pandas as pd

# List of ASX ticker symbols
company_names = ["Telstra", "Woolworths", "Commonwealth Bank"]

# Fields to retrieve
fields = ["Name", "DividendYield", "MarketCap"]

# Function to fetch data for a single company
def fetch_company_data(name):
    try:
        response = mstarpy.search_stock(
            term=name,
            field=fields,
            exchange="XASX",
            pageSize=1
        )
        print(f"API Response for {name}: {response}")  # Debugging
        if response:
            return response[0]  # Return first result
        else:
            print(f"No data found for {name}")
            return None
    except Exception as e:
        print(f"Error fetching data for {name}: {e}")
        return None

# Fetch and display data
for name in company_names:
    data = fetch_company_data(name)
    print(data)
