import mstarpy
import pandas as pd

# Load the file containing company names
file_path = 'companies-list.csv'  # Adjust this path if necessary
company_data = pd.read_csv(file_path)

# Clean the 'Company' column to remove the code in parentheses and "Limited" or "Ltd"
company_data['CleanedCompany'] = (
    company_data['Company']
    .str.split('(').str[0]               # Remove everything after '('
    .str.replace(r'\b(Limited|Ltd)\b', '', regex=True)  # Remove "Limited" or "Ltd"
    .str.strip()                        # Remove extra whitespace
)

# Extract the cleaned company names
company_names = company_data['CleanedCompany'].tolist()

# Fields to retrieve directly
fields = [
    "Name",               # Stock name
    "ROATTM",             # Return on Assets (ROA, trailing twelve months)
    "ROETTM",             # Return on Equity (ROE, trailing twelve months)
    "DividendYield",      # Dividend yield (DIV)
    "DebtEquityRatio",    # Debt-to-Equity Ratio (proxy for DAR)
    "MarketCap",          # Market capitalization (for SIZE)
    "NetIncome"           # Net income
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
                "Net Income (Million)": round(data.get("NetIncome", 0), 2) if data.get("NetIncome") is not None else "Missing"
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
    output_file = "Company_Data.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")
else:
    print("No data was collected.")
