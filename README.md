# QuickCart Regional Sales Performance Dashboard

An end to end Excel project where I took a messy, unusable order export and turned it into a clean, live dashboard that regional managers could actually use to make decisions.

## Tech & Skills Used

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Cleaning](https://img.shields.io/badge/Data%20Cleaning-0078D4?style=for-the-badge)
![Pivot Tables](https://img.shields.io/badge/Pivot%20Tables-4CAF50?style=for-the-badge)
![XLOOKUP](https://img.shields.io/badge/XLOOKUP%20%2F%20VLOOKUP-FF6F00?style=for-the-badge)
![INDEX MATCH](https://img.shields.io/badge/INDEX%20%2B%20MATCH-9C27B0?style=for-the-badge)
![Data Validation](https://img.shields.io/badge/Data%20Validation-00897B?style=for-the-badge)
![Dashboard Design](https://img.shields.io/badge/Dashboard%20Design-E91E63?style=for-the-badge)

## What This Project Is About

I got a raw weekly sales export that honestly looked unusable at first glance. No revenue column filled in, region names spelled five different ways, duplicate rows, blank rows, prices sometimes stored as text with a currency symbol attached. Customer and product info was just a bunch of ID codes, nothing a manager could actually read.

My job in this project was to take that mess and turn it into something a business person could open, filter, and immediately understand. Below is a walk through of every problem I ran into and exactly how I solved it, written in plain language rather than textbook definitions.

---

## Problem 1: The data could not be trusted yet

**What was actually wrong**
The sheet had 40 completely empty rows sitting randomly in between real data, and 186 rows that were exact duplicates of each other. On top of that, the Region, Payment Method, and Order Status columns had the same value typed in five or six different ways ("North", "north", "NORTH", "North " with a trailing space, even "N." as a shortcut).

**How I tackled it**
I never touched the original file directly, I worked on a copy so I always had the messy version as proof of the "before" state. I filtered for blank rows and deleted them. Then I used Excel's Remove Duplicates tool checking every column, which cleared out the 186 repeated rows. For the messy text columns, TRIM and PROPER alone were not enough since some entries were abbreviations or genuine spelling mistakes, so I layered in some IF logic to catch those specific cases too and map everything back to one clean, correct value.

---

## Problem 2: Nobody could tell what the actual revenue was

**What was actually wrong**
The Total Amount column was completely empty for every single row. Worse, Quantity and Unit Price themselves were not reliable either. Some were missing, some had text like "Rs." stuck in front of the number, and a few even had negative values from what looked like typing mistakes.

**How I tackled it**
I built cleaned versions of Quantity and Unit Price first, stripping out currency text and forcing everything into real numbers, and fixing the negative values using ABS. Then I calculated Total Amount as Quantity times Unit Price, but wrapped the whole thing so that if either input was missing, it would not throw an error or silently skip that row. It would show 0 and also get flagged in a separate column as "Incomplete" so I never lost sight of how much of the total was based on real numbers versus how much was missing data.

---

## Problem 3: The data was just codes, nothing a manager could read

**What was actually wrong**
Every order only referenced a Customer ID like "C1572" and a combined product code like "SKU P1002 ELE Pro". None of that means anything to someone trying to understand which customers or products actually drive revenue.

**How I tackled it**
First I pulled apart the product code into its individual pieces (Product ID, category, variant) using text functions, since the code always followed the same pattern. Then I used XLOOKUP to bring in the actual product name, category, and price from the Products sheet. For customer info, I deliberately used INDEX plus MATCH instead, mainly to show I understand both approaches and when each one makes sense, since INDEX plus MATCH works in older Excel versions and does not break if columns get rearranged later. Any row with a missing Customer ID got a clear "No Customer ID" label instead of a confusing error, so that gap stayed visible rather than hidden.

---

## Problem 4: The same mess would just come back next week

**What was actually wrong**
Cleaning the data once does not stop someone from typing "north" again by hand next week when new orders come in.

**How I tackled it**
I built a small Reference List sheet holding the correct spelling for every Region, Payment Method, and Status. Then I applied Data Validation dropdowns on those columns, pointing at that reference list, so going forward, whoever enters new data can only pick from the approved list instead of free typing something inconsistent.

---

## Problem 5: There was no live way to actually look at this data

**What was actually wrong**
Someone had to manually rebuild a summary report every single week for regional managers, which is slow and easy to mess up.

**How I tackled it**
I converted the cleaned data into a proper Excel Table so it grows automatically as new rows get added. Then I built a Pivot Table breaking revenue and order count down by Region and Status, added Slicers for Region, Status, and Loyalty Tier so anyone can filter with one click, and linked a Pivot Chart to it so the visual updates live too. On top I added four KPI cards summarizing Total Revenue, Total Orders, Cancellation Rate, and a Data Completeness Rate, that last one specifically to be upfront about how much of the data had missing customer info instead of quietly hiding it.

---

## Final Numbers

| Metric | Value |
|---|---|
| Total Revenue | Rs 4,27,70,768 |
| Total Orders | 5963 |
| Cancellation Rate | 11.79% |
| Data Completeness Rate | 97.87% |
| Duplicate Rows Removed | 186 |
| Blank Rows Removed | 40 |

## Workbook Structure

| Sheet | What It Holds |
|---|---|
| Orders | Main cleaned, calculated, and enriched transaction data |
| Customers | Clean customer master list |
| Products | Clean product master list |
| Reference_Lists | Approved dropdown values used for Data Validation |
| Dashboard_Data | The Pivot Table behind the dashboard |
| Dashboard | Final KPI cards, chart, and slicers |

## Repo Contents
This repo keeps things simple and focused on the actual project.

QuickCart_Sales_RAW.xlsx, the original messy data before any cleaning, kept as proof of the starting point.

QuickCart_Sales_Project.xlsx, the final cleaned, calculated, enriched, and dashboard version.

screenshots folder, images of the raw data, the cleaning process, and the final dashboard.

README.md, this file.

## Screenshots

### Final dashboard
![Final Dashboard](screenshots/dashboard_final.png)

### Pivot table detail
![Pivot Table Detail](screenshots/pivot_table_detail.png)

## About The Dataset (Full Disclosure)

The dataset used in this project is AI generated synthetic data, not real transaction data from an actual company. I want to be upfront about this rather than let anyone assume otherwise.

It was created using Python along with the Faker library, with realistic Indian names, cities, and product categories, and a fixed random seed so the results are reproducible. The messiness in the data (duplicate rows, blank rows, inconsistent text, missing values, text stored numbers) was deliberately injected on purpose, to simulate the kind of real world data quality problems an analyst actually runs into on the job.

The dataset includes about 6150 order records, 600 customer records, and 180 product records spread across three connected tables, built specifically so I could practice cleaning, validation, lookups across multiple tables, and dashboard building all within one coherent project rather than several disconnected exercises.

Every formula, cleaning decision, and dashboard design choice in this project is my own work, built and verified manually in Excel on top of this dataset.
