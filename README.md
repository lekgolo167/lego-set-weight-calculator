# LEGO Data Analysis and Weight Calculation Tool

This Python project allows users to analyze and calculate details about LEGO sets, including prices (adjusted for inflation), part weights, and price per piece. The data is fetched and parsed from BrickLink and BrickSet websites, organized into a spreadsheet for further analysis.
---

## Features

### 1. **Retrieve LEGO Theme Information**
- Extracts available years for a given LEGO theme from BrickSet.

### 2. **Fetch and Cache LEGO Set Details**
- Retrieves data for LEGO sets within a specific theme and year.
- Caches the data locally to avoid repeated downloads.

### 3. **Calculate Set Weights**
- Computes the total weight of a LEGO set using individual part weights, fetched or cached locally.

### 4. **Spreadsheet Generation**
- Outputs LEGO set data, including:
  - Year, set ID, name, MSRP, inflation-adjusted price, number of pieces, weight, cost per gram, and cost per piece.
- Saves the data in an Excel workbook for analysis.

### 5. **Support for Minifigures and Parts**
- Differentiates between regular parts and minifigures.
- Calculates weights for both categories separately.

### 6. **Directory Management**
- Automatically creates required directory structures for storing data and caching results.

---

## Dependencies

- `sys` and `os`: For system operations and file management.
- `math.floor`: For rounding calculations.
- `requests`: To fetch data from websites.
- `pickle`: For caching part weights locally.
- `bs4.BeautifulSoup`: For HTML parsing.
- `re`: For regular expressions in data extraction.
- `openpyxl`: For Excel workbook management.

---

## Installation

1. Install Python 3.8 or later.
2. Install required packages using pip:
   ```bash
   pip install requests beautifulsoup4 openpyxl
   ```

---

## How to Use

### 1. **Set the Theme**
- Modify the `theme` variable in the `__main__` section to specify the LEGO theme of interest.

### 2. **Run the Script**
```bash
python brickset.py
```

### 3. **Output**
- The script creates an Excel file (`Lego_sets.xlsx`) in the current directory containing all analyzed data.
- It also saves cached data in the `./parts` and `./themes` directories for reuse.

---

## Key Functions

### 1. **`get_theme_years(theme)`**
Fetches the available years for a specified theme from BrickSet.

### 2. **`get_sets(theme, year)`**
Fetches all LEGO sets within a theme and year, including piece counts, prices, and IDs.

### 3. **`parse_set(set_id)`**
Parses a LEGO set's parts and quantities from BrickLink.

### 4. **`get_part_weight(part_id)`**
Calculates the weight of an individual part, with specific handling for minifigures.

### 5. **`get_set_weight(part_weights, set_parts)`**
Computes the total weight of a LEGO set using cached or fetched part weights.

### 6. **`fillout_workbook(filename, sets_year, theme)`**
Populates an Excel workbook with LEGO set data for a specific theme.

### 7. **`create_directories(theme)`**
Ensures required directory structures are created before running the script.

---

## Caching

- Part weights and HTML pages for sets/themes are cached locally to reduce redundant requests to the BrickLink and BrickSet websites.
- Caching improves performance for large datasets and ensures offline availability.

---

## Limitations

1. **API Restrictions**:
   - The tool relies on web scraping, which may break if the structure of the source websites changes.
2. **Incomplete Data Handling**:
   - Sets without prices or weights are logged but may not be included in calculations.

---

## Future Enhancements

1. **Add Support for More Themes**
   - Automatically fetch all themes and process them in bulk.

2. **Improve Error Handling**
   - Handle cases where data is missing more gracefully.

3. **Dynamic Inflation Calculation**
   - Update inflation data dynamically from external sources.

---

## Author
This tool is designed for LEGO enthusiasts and researchers interested in analyzing LEGO set data, parts, and pricing trends. It is a hobbyist project and not affiliated with LEGO or BrickLink.

