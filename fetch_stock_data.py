import mstarpy
import pandas as pd

# List of ASX company names
company_names = ["Telstra", "Woolworths", "Commonwealth Bank"]

# Fields to retrieve directly
fields = [
    "Name",               # Stock name
    "ROATTM",             # Return on Assets (ROA, trailing twelve months)
    "ROETTM",             # Return on Equity (ROE, trailing twelve months)
    "DividendYield",      # Dividend yield (DIV)
    "DebtEquityRatio",    # Debt-to-Equity Ratio (proxy for DAR)
    "MarketCap",          # Market capitalization (for SIZE)
]

# Function to fetch and process company data
def fetch_company_data(name):
    try:
        # Fetch data from Morningstar
        response = mstarpy.search_stock(
            term=name,
            field=fields,
            exchange="XASX",  # ASX Exchange code
            pageSize=1        # Return only the top result
        )
        if response:
            # Extract data and format it
            data = response[0]
            return {
                "Name": data.get("Name", "Unknown"),
                "ROA": round(data.get("ROATTM", 0), 2) if data.get("ROATTM") is not None else "Missing",
                "ROE": round(data.get("ROETTM", 0), 2) if data.get("ROETTM") is not None else "Missing",
                "DIV": round(data.get("DividendYield", 0), 2) if data.get("DividendYield") is not None else "Missing",
                "DAR (Proxy)": round(data.get("DebtEquityRatio", 0), 2) if data.get("DebtEquityRatio") is not None else "Missing",
                "SIZE (Billion)": round(data.get("MarketCap", 0) / 1_000_000_000, 2) if data.get("MarketCap") is not None else "Missing",
            }
        else:
            print(f"No data found for {name}")
            return None
    except Exception as e:
        print(f"Error fetching data for {name}: {e}")
        return None

# Fetch data for all companies
data = []
for name in company_names:
    company_data = fetch_company_data(name)
    if company_data:
        data.append(company_data)

# Convert collected data to a DataFrame
if data:
    df = pd.DataFrame(data)

    # Save the results to an Excel file
    output_file = "ASX_Company_Data.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")
else:
    print("No data was collected.")
