# Sell In Report Generator

This script processes sales data to generate a detailed "Sell In" report. It uses data from multiple sheets in an Excel file, merges and processes the data, and then outputs a formatted report in Excel format.

## Features

- **Data Processing**: Filters and groups sales data based on specific criteria (e.g., sales type, emission date).
- **Data Merging**: Combines data from multiple sheets, including customer locations, item details, and zonification.
- **Report Generation**: Outputs the processed data into a new Excel file, with customized formatting for better readability.

## Requirements

- Python 3.x
- pandas
- openpyxl

```bash
pip install pandas openpyxl
```

## Usage

1. Place the sell_in_mod.xlsx file in the same directory as this script.

2. Run the script:

```bash
python sell_in.py
```

3. The script will generate a new Excel file named Reporte de Sell in (Actualizable).xlsx in the same directory.

## Detailed Workflow

1. Loading Data:

The script loads data from different sheets (LIMA, Listas-clientes, Listas-ubicaciones, and Listas-articulos) in the Excel file.

2. Data Filtering:

Filters the data to include only specific sales types (NC_Ventas, Ventas) and records from the year 2023 onward.
Converts certain columns to uppercase for uniformity.

3. Data Grouping:

Groups the data by specific columns and sums relevant fields.
Fills missing values in ProvinciaDespacho based on grouped CodCliente.

4. Data Merging:

Merges grouped data with additional information from ubicaciones, zonificacion, and articulos.
Fills missing values with "Otros" when necessary.

5. Report Generation:

The final processed data is saved into a new Excel file.

## License

This project is licensed under the MIT License. [See here](https://opensource.org/licenses/MIT) for more details.

## Author

Renato Perez