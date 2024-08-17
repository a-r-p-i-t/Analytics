## Store Inventory: Trends Analysis
This project involves analyzing store-wise data to understand trends, replenishment rates, consumption rates, and average shelf fullness for different brands. The analysis is carried out using Python and various libraries like pandas, matplotlib, seaborn, and xlwings.

## Table of Contents

- [Project Overview](#project-overview)
- [Installation](#installation)
- [Usage](#usage)
- [Dependencies](#dependencies)
- [Key Functionalities](#key-functionalities)
- [Results and Conclusion](#results-and-conclusion)


## Project Overview

The main objectives of this analytics project are:

1. **Data Extraction**: The project starts by extracting data from an Excel file located in the specified directory using the pandas library.

2. **Average Fullness Calculation**: The `avg_fullness` function calculates the average fullness for a given date and brand. It iterates through the data to calculate the average fullness based on the provided date and brand name.

3. **Critical Points Identification**: The `critical_pts` function identifies critical points in the data. It processes the data to determine significant changes in fullness that indicate potential replenishment or consumption events.

4. **Data Visualization**: The project includes data visualization using matplotlib and seaborn libraries. Various charts and graphs are generated to visualize trends, critical points, and rates of replenishment and consumption.

5. **Excel Report Generation**: The analysis results are presented in an Excel report. Dataframes containing MTTR, MTTC, replenishment rate, consumption rate, and average shelf fullness are created and saved in an Excel file.
   

## Installation

1. **Install the required modules:**

    ```bash
    pip install pandas numpy matplotlib xlwings

## Usage

1. **Input Requirements:**
   - Store names (separated by semicolons).
   - Start and end date for the analysis period.

2. **Output:**
   - An Excel workbook titled chart.xlsx containing
       - Data tables summarizing the analysis.
       - Various charts embedded within the workbook for visual representation of the data.
    
3. **Execution:**
   - Run the script **[store.py](./store.py)** and provide the required inputs when prompted.
   - The script will process the data and generate an Excel file with the results.
  
# Dependencies
   
  - **pandas**: For data manipulation.
  - **numpy**: For numerical operations.
  - **matplotlib**: For creating charts.
  - **xlwings**: For interaction with Excel.
  - **datetime**, **re**, **os**: For handling dates, regular expressions, and file system paths, respectively.

# Key Functionalities
1. **Data Extraction:**
    - file_extract(sheet_name): Reads an Excel file for the specified sheet, which contains the store data.

2. **Data Analysis Functions:**
     - **avg_fullness(date, sheet_name):** Computes the average fullness of the store shelves for a specific date.
     - **critical_pts(sheet_name):** Analyzes the inventory data to find critical points where there were significant changes in shelf fullness, calculates metrics like Mean Time to Replenish (MTTR) and Mean Time to Consume (MTTC), and generates various charts to visualize these metrics.

3. **Chart Generation:**
     - Pie charts to show the distribution of replenishment and consumption rates across different days of the week.
     - Bar charts to compare average replenishment and consumption rates by day.
     - Scatter and line plots to visualize critical changes in shelf fullness over time.
     - The charts are added to an Excel sheet using xlwings.

4. **User Interaction:**
     - Collects user input for the store names and date range to analyze.






# Results and Conclusion

This analytics project provides insights into the trends, replenishment rates, consumption rates, and average shelf fullness for different brands in different stores. By analyzing critical points, the project helps in understanding when and how replenishment and consumption events occur. The visualizations provide a clear representation of the data patterns and trends.<br>
<br>



![Chart Title](https://res.cloudinary.com/dqwly03el/image/upload/v1693388698/visualisation_charts_lu5p0c.png)








