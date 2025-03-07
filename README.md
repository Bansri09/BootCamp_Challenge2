# **VBA Challenge - Stock Market Data Analysis**

## **Overview**

This challenge implements a **VBA (Visual Basic for Applications)** script to automate stock market data analysis. The script processes quarterly stock data to calculate various performance metrics such as percentage changes, total volume, and highlights the top performers. It also applies **conditional formatting** to help visualize changes in stock performance.

---

## **Challenge Structure**

The challenge is organized into the following key components:

- **VBA Script Files:** Contains the logic for analyzing and processing stock data.
- **Test File (`alphabetical_testing.xlsx`):** A smaller dataset for quicker testing and development.
- **Results (Screenshots):** Screenshots showing the output of the script when applied to the data.
  
Result of stock analysis results:

<img width="718" alt="Result1" src="https://github.com/user-attachments/assets/a7594d38-48b4-4676-8366-fe08a0c7ab5c" />

Result of conditional formatting applied to the data:

<img width="898" alt="Result2" src="https://github.com/user-attachments/assets/77e89683-3db4-4301-9b0e-f2c7d0f43044" />

---

## **Features**

### 1. **Data Retrieval & Analysis**

The VBA script loops through stock data for each quarter and outputs:

- **Ticker Symbol**
- **Quarterly Change** (The change from the opening to closing price)
- **Percentage Change** (Percentage change from opening to closing price)
- **Total Volume** of stocks traded during the quarter

### 2. **Performance Highlights**

The script calculates and outputs the following:

- **Greatest Percentage Increase**
- **Greatest Percentage Decrease**
- **Greatest Total Volume**

### 3. **Conditional Formatting**

The script applies conditional formatting to the following columns:

- **Quarterly Change:** Green for positive values, Red for negative values.
- **Percentage Change:** Green for positive values, Red for negative values.

### 4. **Looping Across Worksheets**

The script can be run across all worksheets in the workbook, where each worksheet represents data for a different quarter.

---




