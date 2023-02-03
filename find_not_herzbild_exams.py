import os
import pandas as pd

# Load the data from the excel file
try:
    xlsx_file = "Rechnungsliste_HerzBild.xlsx"
    df_xlsx = pd.read_excel(xlsx_file, skiprows=1)
except FileNotFoundError:
    print(f"Error: The file {xlsx_file} was not found.")
except Exception as e:
    print(f"Error: An error occurred while reading the excel file: {e}")

# Load the data from the csv file
csv_file = None
for file in os.listdir("Adresslisten"):
    if file.endswith(".csv"):
        csv_file = f"Adresslisten/{file}"
        break

if csv_file is None:
    print("Error: No CSV file was found in the current directory.")
else:
    try:
        df_csv = pd.read_csv(csv_file, skiprows=9, encoding='ISO-8859-1', sep=";")
        df_csv = df_csv.rename(columns={"Patient-ID": "PatientenID"})
    except FileNotFoundError:
        print(f"Error: The file {csv_file} was not found.")
    except Exception as e:
        print(f"Error: An error occurred while reading the csv file: {e}")

# Compare the data
try:
    # Get the "PatientenID" column from both data frames
    xlsx_ids = df_xlsx["PatientenID"]
    csv_ids = df_csv["PatientenID"]

    # Find the entries in the csv file that are not in the excel file
    missing_entries = df_csv[~df_csv["PatientenID"].isin(xlsx_ids)]

    # Filter the missing entries to only include rows with "Keine" or "Nan" in the "Abrechnungsart" column
    missing_entries = missing_entries[missing_entries["Abrechnungsart"].isin(["Keine", "Nan"])]

    # Write the filtered entries to a csv file
    missing_entries.to_csv("output.csv", index=False, sep=";")
    print("The filtered entries have been written to 'output.csv'.")
except KeyError as e:
    print(f"Error: The column 'PatientenID' or 'Abrechnungsart' was not found in one of the data frames: {e}")
except Exception as e:
    print(f"Error: An error occurred while comparing the data: {e}")
