import mstarpy
import pandas as pd
import datetime

# List of ASX company names
company_names = ["Telstra", "Woolworths", "Commonwealth Bank"]

# Data Point IDs for key financial variables
data_point_ids = {
    "ROA": "IFBS002270",  # Total Assets (proxy for ROA calculation)
    "ROE": "IFBS002220",  # Total Equity (proxy for ROE calculation)
    "NetIncome": "IFIS001100",  # Net Income
    "DebtEquityRatio": "IFBS002646",  # Debt-to-Equity Ratio
    "DividendYield": "SAL0000001",  # Dividend Yield
    "MarketCap": "IFBS001170"  # Market Capitalization
}

# Helper function to extract multi-year data
def extract_multi_year_data(rows, data_point_id, years):
    for row in rows:
        if row["dataPointId"] == data_point_id:
            return {year: row["datum"][i] if i < len(row["datum"]) else None for i, year in enumerate(years)}
    return {year: None for year in years}

# Function to fetch data for a single company
def fetch_company_data(name):
    try:
        response = mstarpy.search_stock(
            term=name,
            field=list(data_point_ids.values()),
            exchange="XASX",
            pageSize=1
        )
        if response:
            financial_data = response[0]
            rows = financial_data.get("rows", [])
            column_defs = financial_data.get("columnDefs", [])
            
            # Extract years from column definitions
            years = [datetime.datetime.strptime(col, "%Y%m%d").year for col in column_defs]
            
            # Extract data for each year
            result = {"Name": name}
            for var_name, data_point_id in data_point_ids.items():
                result[var_name] = extract_multi_year_data(rows, data_point_id, years)
            
            return result
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

# Convert collected data to multiple DataFrames (one per year)
if data:
    # Extract all years
    all_years = set()
    for company in data:
        for year in company["ROA"].keys():
            all_years.add(year)
    all_years = sorted(all_years)

    # Create a DataFrame for each year
    for year in all_years:
        rows = []
        for company in data:
            row = {"Name": company["Name"], "Year": year}
            for var, yearly_data in company.items():
                if var != "Name":
                    row[var] = yearly_data[year]
            rows.append(row)

        # Convert rows to a DataFrame
        df = pd.DataFrame(rows)

        # Save each year's data to a separate Excel sheet
        output_file = f"ASX_Company_Data_{year}.xlsx"
        df.to_excel(output_file, index=False)
        print(f"Data for {year} saved to {output_file}")
else:
    print("No data was collected.")
