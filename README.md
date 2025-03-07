# **VBA Challenge - Stock Market Data Analysis**

## **Overview**

This challenge implements a **VBA (Visual Basic for Applications)** script to automate stock market data analysis. The script processes quarterly stock data to calculate various performance metrics such as percentage changes, total volume, and highlights the top performers. It also applies **conditional formatting** to help visualize changes in stock performance.

---

## **Challenge Structure**

The challenge is organized into the following key components:

- **VBA Script Files:** Contains the logic for analyzing and processing stock data.
- **Test File (`alphabetical_testing.xlsx`):** A smaller dataset for quicker testing and development.
- **Results (Screenshots):** Screenshots showing the output of the script when applied to the data.

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

## **How to Use**

1. **Clone the Repository:**
   Clone the repository to your local machine using:
   ```bash
   git clone https://github.com/yourusername/VBA-challenge.git
