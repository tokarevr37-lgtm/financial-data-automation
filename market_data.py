import csv
from pathlib import Path

import requests
from openpyxl import Workbook


API_URL = "https://financialmodelingprep.com/api/v3/quote/AAPL,MSFT,GOOGL,AMZN,TSLA?apikey=demo"

CSV_FILE = "financial_data.csv"
EXCEL_FILE = "financial_data.xlsx"


def fetch_market_data():
    try:
        response = requests.get(API_URL, timeout=10)
        response.raise_for_status()
        data = response.json()

        formatted_data = []

        for item in data:
            formatted_data.append({
                "symbol": item["symbol"],
                "price": item["price"],
                "change_percent": item["changesPercentage"],
                "day_high": item["dayHigh"],
                "day_low": item["dayLow"],
                "volume": item["volume"],
            })

        return formatted_data

    except requests.RequestException as error:
        print(f"API request failed: {error}")
        return []


def save_to_csv(data, filename):
    with open(filename, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)


def save_to_excel(data, filename):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Financial Data"

    headers = list(data[0].keys())
    sheet.append(headers)

    for row in data:
        sheet.append(list(row.values()))

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        sheet.column_dimensions[column_letter].width = max_length + 2

    workbook.save(filename)


def main():
    data = fetch_market_data()

    if not data:
        print("No data collected.")
        return

    save_to_csv(data, CSV_FILE)
    save_to_excel(data, EXCEL_FILE)

    print(f"Done. Exported {len(data)} records.")
    print(f"CSV saved to: {Path(CSV_FILE).resolve()}")
    print(f"Excel saved to: {Path(EXCEL_FILE).resolve()}")


if __name__ == "__main__":
    main()
